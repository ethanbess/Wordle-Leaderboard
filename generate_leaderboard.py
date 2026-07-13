"""
Reads games.csv (produced by export_games.ps1 from the Wordle Excel file)
and generates leaderboard.json. Run once per day via Windows Task Scheduler -
see daily_update.ps1, which runs export_games.ps1 first.
"""
import csv
import json
import statistics
from datetime import date, datetime, timedelta
from pathlib import Path

CSV_PATH = Path(__file__).parent / "games.csv"
OUTPUT_PATH = Path(__file__).parent / "leaderboard.json"

ROLLING_WINDOW_DAYS = 21
STREAK_CAP = 21
STREAK_WEIGHT = 0.06
MISSED_DAY_PENALTY = 0.03
MIN_GAMES_IN_WINDOW = 5


def load_games():
    games = []
    with CSV_PATH.open(newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            d = datetime.strptime(row["Date"], "%Y-%m-%d").date()
            games.append((d, row["Player"].strip(), int(row["Guesses"])))
    return games


def compute_leaderboard(games, today=None):
    today = today or date.today()
    window_start = today - timedelta(days=ROLLING_WINDOW_DAYS - 1)

    by_player = {}
    for d, player, guesses in games:
        by_player.setdefault(player, []).append((d, guesses))

    results = []
    for player, entries in by_player.items():
        # 21-day rolling average
        in_window = [g for d, g in entries if window_start <= d <= today]
        if len(in_window) < MIN_GAMES_IN_WINDOW:
            continue  # not enough games in the window to qualify for ranking
        avg = statistics.mean(in_window)

        # Played streak: consecutive calendar days with >=1 game, ending today or yesterday
        played_days = {d for d, _ in entries}
        last_played = max(played_days)
        if (today - last_played).days > 1:
            streak = 0
        else:
            streak = 0
            cursor = last_played
            while cursor in played_days:
                streak += 1
                cursor -= timedelta(days=1)

        # Missed-day penalty: days with no game, within the rolling window but
        # never before the player's first-ever logged game (don't penalize
        # newcomers for weeks before they joined).
        first_played = min(played_days)
        penalty_start = max(window_start, first_played)
        days_in_scope = (today - penalty_start).days + 1
        days_played_in_scope = len({d for d in played_days if penalty_start <= d <= today})
        missed_days = days_in_scope - days_played_in_scope
        penalty = missed_days * MISSED_DAY_PENALTY

        score = avg + penalty - (min(streak, STREAK_CAP) * STREAK_WEIGHT)

        results.append({
            "player": player,
            "score": round(score, 3),
            "streak": streak,
            "avg": round(avg, 2),
        })

    results.sort(key=lambda r: r["score"])
    for i, r in enumerate(results, start=1):
        r["rank"] = i

    return results


def main():
    games = load_games()
    leaderboard = compute_leaderboard(games)
    payload = {
        "generated_at": date.today().isoformat(),
        "players": leaderboard,
    }
    OUTPUT_PATH.write_text(json.dumps(payload, indent=2))
    print(f"Wrote {len(leaderboard)} players to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
