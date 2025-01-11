import requests
import pandas as pd
import time
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

API_URL = "https://graphql.anilist.co"

def fetch_anime_data(username, season, year):
    query = """
    query ($username: String, $season: MediaSeason, $year: Int) {
      MediaListCollection(userName: $username, type: ANIME) {
        lists {
          entries {
            media {
              title {
                romaji
              }
              season
              seasonYear
              averageScore
              popularity
            }
            score
            status
          }
        }
      }
      Media(season: $season, seasonYear: $year, type: ANIME) {
        title {
          romaji
        }
        season
        seasonYear
        averageScore
        popularity
      }
    }
    """
    variables = {
        "username": username,
        "season": season,
        "year": year
    }

    response = requests.post(API_URL, json={"query": query, "variables": variables})
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error: {response.status_code}")
        print(response.json())
        return None

def calculate_combined_weighted_mean(scores, popularities, season_mean, global_mean, seen_count, total_count, wp=0.6, wa=0.4):
    # Popularitäts-basierter gewichteter Mittelwert
    popularity_weighted_mean = (
        sum(score * pop for score, pop in zip(scores, popularities) if score > 0) /
        sum(pop for score, pop in zip(scores, popularities) if score > 0)
        if popularities else 0
    )

    # Aktivitäts-basierter gewichteter Mittelwert
    activity_ratio = seen_count / total_count if total_count > 0 else 0
    activity_weighted_mean = (
        activity_ratio * season_mean + (1 - activity_ratio) * global_mean
    )

    # Kombinierter gewichteter Mittelwert
    combined_weighted_mean = round(
        wp * popularity_weighted_mean + wa * activity_weighted_mean, 2
    )

    return combined_weighted_mean

def calculate_statistics(data, season, year, global_mean=7.0, total_count=50):
    if not data or "data" not in data or "MediaListCollection" not in data["data"]:
        return {
            "season": season,
            "year": year,
            "anime_count": 0,
            "mean_score": 0,
            "weighted_mean": 0,
            "anime_list": []
        }

    anime_list = []
    scores = []
    popularities = []

    for list_data in data["data"]["MediaListCollection"]["lists"]:
        for entry in list_data["entries"]:
            if entry.get("status") == "COMPLETED":
                media = entry["media"]
                if media["season"] == season and media["seasonYear"] == year:
                    score = entry.get("score", 0)
                    title = media["title"]["romaji"]
                    formatted_entry = f"{score} - {title}"
                    anime_list.append(formatted_entry)
                    if score:
                        scores.append(score)
                    popularity = media.get("popularity", 0)
                    popularities.append(popularity)

    mean_score = round(sum(scores) / len(scores), 2) if scores else 0
    weighted_mean = calculate_combined_weighted_mean(
        scores, popularities, season_mean=mean_score, global_mean=global_mean,
        seen_count=len(anime_list), total_count=total_count
    )

    return {
        "season": season,
        "year": year,
        "anime_count": len(anime_list),
        "mean_score": mean_score,
        "weighted_mean": weighted_mean,
        "anime_list": anime_list
    }

def analyze_all_seasons(username, start_year=2006):
    seasons = ["WINTER", "SPRING", "SUMMER", "FALL"]
    current_year = pd.Timestamp.now().year
    all_stats = []
    request_count = 0

    for year in range(start_year, current_year + 1):
        for season in seasons:
            if request_count >= 30:
                print("Rate limit reached. Pausing for 60 seconds...")
                time.sleep(60)
                request_count = 0

            print(f"Fetching data for {season} {year}...")
            data = fetch_anime_data(username, season, year)
            stats = calculate_statistics(data, season, year)
            all_stats.append(stats)
            request_count += 1
            time.sleep(2)

    return all_stats

def save_to_excel_with_formatting(all_stats, filename):
    rows = []
    for stat in all_stats:
        row = {
            "season": stat["season"],
            "year": stat["year"],
            "anime_count": stat["anime_count"],
            "mean_score": stat["mean_score"],
            "weighted_mean": stat["weighted_mean"]
        }
        for i, anime in enumerate(stat["anime_list"]):
            row[f"anime_{i + 1}"] = anime
        rows.append(row)

    df = pd.DataFrame(rows)
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    filepath = f"{current_date}_{filename}.xlsx"
    df.to_excel(filepath, index=False)

    wb = load_workbook(filepath)
    ws = wb.active

    apply_custom_formatting(ws, "C", top=2, bottom=2)
    apply_custom_formatting(ws, "D", top=2, bottom=2)
    apply_custom_formatting(ws, "E", top=2, bottom=2)

    wb.save(filepath)
    print(f"Data saved and formatted to {filepath}")

def apply_custom_formatting(sheet, column_letter, top=2, bottom=2, ignore_zeros=True):
    values = [
        (row, sheet[f"{column_letter}{row}"].value)
        for row in range(2, sheet.max_row + 1)
        if not ignore_zeros or sheet[f"{column_letter}{row}"].value != 0
    ]

    sorted_values = sorted(values, key=lambda x: x[1])
    lowest = sorted_values[:bottom]
    highest = sorted_values[-top:]

    green_fill = PatternFill(start_color="32CD32", end_color="32CD32", fill_type="solid")
    red_fill = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")

    for row, _ in lowest:
        sheet[f"{column_letter}{row}"].fill = red_fill
    for row, _ in highest:
        sheet[f"{column_letter}{row}"].fill = green_fill

if __name__ == "__main__":
    username = input("Please enter your AniList username: ").strip()
    if not username:
        print("No username entered. Exiting program.")
    else:
        stats = analyze_all_seasons(username)
        save_to_excel_with_formatting(stats, "anime_season_stats")
