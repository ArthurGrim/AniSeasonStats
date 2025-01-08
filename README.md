# AniSeasonStats
AniSeasonStats

AniSeasonStats is a Python program designed to analyze a user's seasonal anime watching habits on AniList. By leveraging the AniList API, the program retrieves data about anime watched during each season (Winter, Spring, Summer, Fall) from a specified start year and calculates various statistics. The results are exported into an Excel file with conditional formatting for better visualization.
Features

    Retrieves anime data for a specified AniList username using the AniList API.
    Calculates key statistics for each season and year, including:
        Total number of anime watched per season.
        Mean score of anime watched.
        Weighted mean score based on popularity.
    Highlights the top 2 and bottom 2 values in each column (ignoring zero values) using conditional formatting in the generated Excel file.
    Ensures accurate calculations by excluding scores of 0 and avoids division errors.

Installation

    Clone the repository to your local machine:

git clone https://github.com/<YourUsername>/AniSeasonStats.git

Navigate to the project folder:

cd AniSeasonStats

Create and activate a virtual environment:

python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

Install the required dependencies:

    pip install -r requirements.txt

Usage

    Run the program:

    python AniSeasonStats.py

    Enter your AniList username when prompted.

    The program will:
        Fetch and analyze your anime watching data for each season from 2006 to the current year.
        Generate an Excel file named with the current date and save it in the project directory.

    Open the Excel file to explore the statistics.

Output Format

    The Excel file includes the following columns:
        Season: The anime season (Winter, Spring, Summer, Fall).
        Year: The corresponding year.
        Anime Count: Number of anime watched during that season.
        Mean Score: Average score of the anime watched.
        Weighted Mean: Weighted average score based on popularity.
        Anime List: List of anime watched in the format: Score - Title.

Dependencies

The program relies on the following Python packages:

    pandas: For data manipulation and exporting to Excel.
    requests: For fetching data from the AniList API.
    openpyxl: For creating and modifying Excel files.

Notes

    The program respects AniList's API rate limit of 30 requests per minute by including delays between API calls.
    Conditional formatting highlights the top 2 and bottom 2 values in Anime Count, Mean Score, and Weighted Mean columns for better visual insights.