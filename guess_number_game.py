import random


class GuessNumberGame:
    def __init__(self, target=None, lower=1, upper=100, max_attempts=5):
        self.lower = lower
        self.upper = upper
        self.max_attempts = max_attempts
        self.target = target if target is not None else random.randint(lower, upper)
        self.attempts = 0
        self.finished = False

    def guess(self, value):
        if self.finished:
            return {"status": "finished", "message": "遊戲已結束，請重新開始新的一局。"}

        if value < self.lower or value > self.upper:
            return {
                "status": "invalid",
                "message": f"請輸入 {self.lower} 到 {self.upper} 之間的數字。",
            }

        self.attempts += 1

        if value == self.target:
            self.finished = True
            return {
                "status": "correct",
                "message": f"恭喜你猜中了！答案就是 {self.target}。",
            }

        if self.attempts >= self.max_attempts:
            self.finished = True
            return {
                "status": "game_over",
                "message": f"次數用完了，答案是 {self.target}。",
            }

        if value < self.target:
            return {"status": "low", "message": "太小了，再試一次。"}

        return {"status": "high", "message": "太大了，再試一次。"}


def main():
    game = GuessNumberGame()
    print(f"歡迎來玩猜數字遊戲！請猜 {game.lower} 到 {game.upper} 之間的數字。")
    print(f"你共有 {game.max_attempts} 次機會。")

    while not game.finished:
        raw = input("請輸入你的猜測：").strip()
        if not raw.isdigit():
            print("請輸入有效的整數。")
            continue

        result = game.guess(int(raw))
        print(result["message"])


if __name__ == "__main__":
    main()
