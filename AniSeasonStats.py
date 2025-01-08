import requests
import pandas as pd
import time
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import ColorScaleRule

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

def calculate_statistics(data, season, year):
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
            if entry.get("status") == "COMPLETED":  # Nur abgeschlossene Anime
                media = entry["media"]
                if media["season"] == season and media["seasonYear"] == year:
                    score = entry.get("score", 0)
                    title = media["title"]["romaji"]
                    formatted_entry = f"{score} - {title}"
                    anime_list.append(formatted_entry)
                    if score:  # Nur gültige Scores
                        scores.append(score)
                    popularity = media.get("popularity", 0)
                    popularities.append(popularity)

    # Berechnung der Statistiken
    mean_score = round(sum(scores) / len(scores), 2) if scores else 0
    weighted_mean = round(
        sum(score * pop for score, pop in zip(scores, popularities) if score > 0) /
        sum(pop for score, pop in zip(scores, popularities) if score > 0),
        2
    ) if popularities else 0

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
            if request_count >= 30:  # Wenn das Rate-Limit erreicht ist, warten
                print("Rate limit reached. Pausing for 60 seconds...")
                time.sleep(60)
                request_count = 0

            print(f"Fetching data for {season} {year}...")
            data = fetch_anime_data(username, season, year)
            stats = calculate_statistics(data, season, year)
            all_stats.append(stats)
            request_count += 1

            # Wartezeit von 2 Sekunden, um das API-Limit einzuhalten
            time.sleep(2)

    return all_stats

# Hauptfunktion für die Speicherung und Formatierung
def save_to_excel_with_formatting(all_stats, filename):
    import pandas as pd
    from datetime import datetime

    # DataFrame erstellen
    rows = []
    for stat in all_stats:
        row = {
            "season": stat["season"],
            "year": stat["year"],
            "anime_count": stat["anime_count"],
            "mean_score": round(stat["mean_score"], 2),
            "weighted_mean": round(stat["weighted_mean"], 2),
        }
        for i, anime in enumerate(stat["anime_list"]):
            row[f"anime_{i + 1}"] = anime
        rows.append(row)

    df = pd.DataFrame(rows)

    # Speichern als Excel
    current_date = datetime.now().strftime("%Y-%m-%d")
    filepath = f"{current_date}_{filename}.xlsx"
    df.to_excel(filepath, index=False)

    # Bedingte Formatierung hinzufügen
    wb = load_workbook(filepath)
    ws = wb.active

    apply_custom_formatting(ws, "C", top=2, bottom=2)  # Anime Count
    apply_custom_formatting(ws, "D", top=2, bottom=2)  # Mean Score
    apply_custom_formatting(ws, "E", top=2, bottom=2)  # Weighted Mean

    wb.save(filepath)
    print(f"Data saved and formatted to {filepath}")

# Funktion zur Anwendung der bedingten Formatierung
def apply_custom_formatting(sheet, column_letter, top=2, bottom=2, ignore_zeros=True):
    values = []
    # Werte aus der Spalte sammeln (außer der Header-Zeile)
    for row in range(2, sheet.max_row + 1):  # Ab der zweiten Zeile
        cell_value = sheet[f"{column_letter}{row}"].value
        if ignore_zeros and (cell_value == 0 or cell_value is None):  # Ignoriere 0-Werte
            continue
        values.append((row, cell_value))
    
    # Werte sortieren nach den Zahlenwerten
    sorted_values = sorted(values, key=lambda x: x[1])

    # Unterste und oberste Werte auswählen
    lowest = sorted_values[:bottom]  # Unterste `bottom` Werte
    highest = sorted_values[-top:]  # Oberste `top` Werte

    # Farben definieren
    green_fill = PatternFill(start_color="32CD32", end_color="32CD32", fill_type="solid")
    red_fill = PatternFill(start_color="FF6347", end_color="FF6347", fill_type="solid")

    # Farben anwenden
    for row, _ in lowest:
        sheet[f"{column_letter}{row}"].fill = red_fill
    for row, _ in highest:
        sheet[f"{column_letter}{row}"].fill = green_fill

if __name__ == "__main__":
    # Nutzer nach dem Username fragen
    username = input("Please enter your AniList username: ").strip()
    
    if not username:
        print("No username entered. Exiting program.")
    else:
        stats = analyze_all_seasons(username)
        save_to_excel_with_formatting(stats, "anime_season_stats")

