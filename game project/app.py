from flask import Flask, jsonify, request, render_template, session
import random
import time
import pandas as pd
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)

ROW_COUNT = 6
COLUMN_COUNT = 7

PLAYER = 1
AI = 2
EMPTY = 0

board = [[EMPTY for _ in range(COLUMN_COUNT)] for _ in range(ROW_COUNT)]
turn = PLAYER

game_history = []

EXCEL_PATH = r'C:\Users\Muhammad Faiq\Desktop\game_history.xlsx'


def save_history_to_excel():
    if game_history:
        df = pd.DataFrame(game_history)
        df.to_excel(EXCEL_PATH, index=False)


def reset_board():
    global board, turn
    board = [[EMPTY for _ in range(COLUMN_COUNT)] for _ in range(ROW_COUNT)]
    turn = PLAYER
    session['moves'] = []


def drop_piece(row, col, piece):
    board[row][col] = piece


def is_valid_location(col):
    return board[ROW_COUNT - 1][col] == EMPTY


def get_next_open_row(col):
    if col < 0 or col >= COLUMN_COUNT:
        return None  # invalid column index
    for r in range(ROW_COUNT):
        if board[r][col] == EMPTY:
            return r
    return None  # column full


def winning_move(piece):
    # Horizontal
    for c in range(COLUMN_COUNT - 3):
        for r in range(ROW_COUNT):
            if all(board[r][c+i] == piece for i in range(4)):
                return True
    # Vertical
    for c in range(COLUMN_COUNT):
        for r in range(ROW_COUNT - 3):
            if all(board[r+i][c] == piece for i in range(4)):
                return True
    # Positive Diagonal
    for c in range(COLUMN_COUNT - 3):
        for r in range(ROW_COUNT - 3):
            if all(board[r+i][c+i] == piece for i in range(4)):
                return True
    # Negative Diagonal
    for c in range(COLUMN_COUNT - 3):
        for r in range(3, ROW_COUNT):
            if all(board[r-i][c+i] == piece for i in range(4)):
                return True
    return False


def get_valid_locations():
    return [col for col in range(COLUMN_COUNT) if is_valid_location(col)]


def can_win_next(col, piece):
    if not is_valid_location(col):
        return False
    row = get_next_open_row(col)
    if row is None:
        return False
    board[row][col] = piece
    win = winning_move(piece)
    board[row][col] = EMPTY  # Undo move
    return win


def pick_best_move(piece):
    valid_locations = get_valid_locations()
    for col in valid_locations:
        if can_win_next(col, piece):
            return col
    opponent = PLAYER if piece == AI else AI
    for col in valid_locations:
        if can_win_next(col, opponent):
            return col
    return random.choice(valid_locations) if valid_locations else None


def make_ai_move():
    col = pick_best_move(AI)
    if col is not None and is_valid_location(col):
        row = get_next_open_row(col)
        if row is not None:  # ✅ Extra safety!
            drop_piece(row, col, AI)
            return row, col
    return None, None


@app.route("/")
def index():
    reset_board()
    return render_template("index.html")


@app.route("/start_game", methods=["POST"])
def start_game():
    player_name = request.json.get("player_name")
    if not player_name or not player_name.strip():
        return jsonify({"status": "error", "message": "Player name is required."})
    session['player_name'] = player_name
    session['ai_name'] = "AI Bot"
    reset_board()
    return jsonify({"status": "ok"})


@app.route("/get_board")
def get_board():
    return jsonify(board, turn)


@app.route("/move", methods=["POST"])
def make_move():
    global turn
    try:
        col = int(request.json.get("column"))  # ✅ Make sure it's an integer
    except:
        return jsonify({"status": "error", "message": "Invalid column input."})

    # ✅ Debug print
    print("===== MOVE REQUEST =====")
    print(f"TURN: {'PLAYER' if turn == PLAYER else 'AI'}")
    print(f"COLUMN REQUESTED: {col}")
    print("BOARD BEFORE MOVE:")
    for row_debug in reversed(board):
        print(row_debug)
    print("========================")

    row = get_next_open_row(col)
    if row is None:
        return jsonify({"status": "error", "message": f"Column {col} is full. Try another."})

    drop_piece(row, col, turn)
    session['moves'].append([int(row), int(col), int(turn)])

    if winning_move(turn):
        winner_name = session['player_name'] if turn == PLAYER else session['ai_name']
        game_history.append({
            'game_number': len(game_history) + 1,
            'player_name': session['player_name'],
            'winner': winner_name
        })
        save_history_to_excel()
        reset_board()
        return jsonify({"status": "win", "winner": turn, "winner_name": winner_name, "board": board})

    if is_board_full():
        game_history.append({
            'game_number': len(game_history) + 1,
            'player_name': session['player_name'],
            'winner': 'Draw'
        })
        save_history_to_excel()
        reset_board()
        return jsonify({"status": "draw", "board": board})

    turn = AI if turn == PLAYER else PLAYER

    if turn == AI:
        time.sleep(1)
        ai_row, ai_col = make_ai_move()
        if ai_row is not None:
            session['moves'].append([int(ai_row), int(ai_col), AI])
            if winning_move(AI):
                winner_name = session['ai_name']
                game_history.append({
                    'game_number': len(game_history) + 1,
                    'player_name': session['player_name'],
                    'winner': winner_name
                })
                save_history_to_excel()
                reset_board()
                return jsonify({"status": "win", "winner": AI, "winner_name": winner_name, "board": board})
        if is_board_full():
            game_history.append({
                'game_number': len(game_history) + 1,
                'player_name': session['player_name'],
                'winner': 'Draw'
            })
            save_history_to_excel()
            reset_board()
            return jsonify({"status": "draw", "board": board})
        turn = PLAYER

    return jsonify({"status": "ok", "board": board, "turn": turn})


@app.route("/history")
def get_history():
    return jsonify(game_history)


def is_board_full():
    return all(board[ROW_COUNT - 1][col] != EMPTY for col in range(COLUMN_COUNT))


if __name__ == "__main__":
    app.run(debug=True)
