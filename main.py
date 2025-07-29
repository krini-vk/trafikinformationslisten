import datetime
import os
import requests
from docxtpl import DocxTemplate
import pandas as pd
from dotenv import load_dotenv


def get_last_log_date():
    """
    Get the last date from the log file.
    Returns None if the log file does not exist or is empty.
    """
    try:
        with open("log.txt", "r") as log_file:
            lines = log_file.readlines()
            if lines:
                last_line = lines[-1].strip()
                return datetime.datetime.fromisoformat(
                    last_line.split("---")[0].strip()
                ).date()
            else:
                return None
    except FileNotFoundError:
        return None


def update_log(message):
    """
    Update the log file with a message.
    """
    for log_path in log_paths:
        with open(log_path, "a") as log_file:
            log_file.write(f"{datetime.datetime.now().isoformat()} --- {message}\n")


def get_data():
    url = "https://kortservice.vejle.dk/gis/rest/services/SEPTIMA/vej_sw/MapServer/23/query"
    params = {
        "returnGeometry": "true",
        "outSR": "25832",
        "f": "geojson",
        "outFields": "*",
        "where": "traficstatus = 'Trafikudmeldt' and dagetilstart < 100",
        "orderByFields": "startdate",
    }

    # Make the request
    response = requests.get(url, params=params)
    data = response.json()
    # Extract name values from features
    porperties = [feature["properties"] for feature in data.get("features", [])]
    return porperties


def format_data(df: pd.DataFrame):
    # create new columns for road info and user direction
    df["road_info"] = df["oov2roaduserdescription"].apply(
        lambda x: x.split("\n")[0].strip() if isinstance(x, str) else None
    )
    df["user_direction"] = df["oov2roaduserdescription"].apply(
        lambda x: " ".join(x.split("\n")[1:]).strip() if isinstance(x, str) else None
    )

    # rename column dataframe columns to match the template
    df.rename(
        columns={
            "oov2roadinfo": "title",
            "contractorcontactperson": "contractor_contact_person",
            "contractormobile": "contractor_phone",
            "ownermailaddress": "contractor_email",
            "name": "contractor_company",
            "oov2roadmarkstart": "starttime",
            "oov2roadmarkend": "endtime",
        },
        inplace=True,
    )

    # add starttime and endtime to period, if startdate and enddate are the same
    df["duration"] = df.apply(
        lambda row: (
            f"{row['startdate']} fra kl. {row['starttime']} til kl. {row['endtime']}"
            if row["startdate"] == row["enddate"]
            and row["starttime"] is not None
            and row["endtime"] is not None
            else f"fra {row['startdate']} til {row['enddate']}"
        ),
        axis=1,
    )

    # convert date columns to datetime
    df["startdate"] = pd.to_datetime(df["startdate"], errors="coerce")
    df["enddate"] = pd.to_datetime(df["enddate"], errors="coerce")


def main():
    last_log_date = get_last_log_date()

    # get the data from the API
    data = get_data()
    df = pd.DataFrame(data)
    format_data(df)

    # read the template and render the data
    doc_template = DocxTemplate("trafik_info_template.docx")

    if last_log_date is not None:
        new_roadwork = df[df["startdate"] > str(last_log_date)].to_dict(
            orient="records"
        )
        old_roadwork = df[df["startdate"] <= str(last_log_date)].to_dict(
            orient="records"
        )
    else:
        new_roadwork = df.to_dict(orient="records")
        old_roadwork = []

    contents = {
        "new_roadwork": new_roadwork,
        "old_roadwork": old_roadwork,
        "today_date": today_date,
        "current_year": today_date.year,
        "week_number": today_date.isocalendar()[1],
    }
    doc_template.render(contents)
    doc_template.save(output_path)


if __name__ == "__main__":
    try:
        # Set up environment
        load_dotenv()  # Load environment variables from .env

        today_date = datetime.datetime.today().date()

        output_folder = os.getenv("OUTPUT_FOLDER")  # Get from .env
        filename = f"trafik informations listen {today_date}.docx"
        output_path = os.path.join(output_folder, filename)

        log_filename = "log.txt"
        log_paths = [os.path.join(output_folder, log_filename), log_filename]

        # Run the main function
        main()
        update_log("Script executed successfully.")
    except Exception as e:
        update_log(f"Error: {str(e)}")
        raise e
