from guess_number_game import GuessNumberGame


def test_reports_too_small_and_too_large_and_correct():
    game = GuessNumberGame(target=42, max_attempts=5)

    low = game.guess(10)
    high = game.guess(90)
    hit = game.guess(42)

    assert low["status"] == "low"
    assert "太小" in low["message"]
    assert high["status"] == "high"
    assert "太大" in high["message"]
    assert hit["status"] == "correct"
    assert "猜中了" in hit["message"]


def test_rejects_out_of_range_guess():
    game = GuessNumberGame(target=50, lower=1, upper=100)

    result = game.guess(101)

    assert result["status"] == "invalid"
    assert "1 到 100" in result["message"]


def test_ends_game_after_max_attempts():
    game = GuessNumberGame(target=7, max_attempts=2)

    first = game.guess(1)
    second = game.guess(2)
    after_end = game.guess(7)

    assert first["status"] == "low"
    assert second["status"] == "game_over"
    assert "答案是 7" in second["message"]
    assert after_end["status"] == "finished"
