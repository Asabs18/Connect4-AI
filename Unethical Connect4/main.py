import numpy as np
import random
import math
from pyautogui import *
import pyautogui
import time
import keyboard
import win32api, win32con

BOARDPOS = [
    [(580, 215), (700, 215), (830, 215), (960, 215), (1080, 215), (1200, 215), (1330, 215)],
    [(580, 350), (700, 350), (830, 350), (960, 350), (1080, 350), (1200, 350), (1330, 350)],
    [(580, 470), (700, 470), (830, 470), (960, 470), (1080, 470), (1200, 470), (1330, 470)],
    [(580, 600), (700, 600), (830, 600), (960, 600), (1080, 600), (1200, 600), (1330, 600)],
    [(580, 720), (700, 720), (830, 720), (960, 720), (1080, 720), (1200, 720), (1330, 720)],
    [(580, 850), (700, 850), (830, 850), (960, 850), (1080, 850), (1200, 850), (1330, 850)]]

BLACK = (0,0,0)
YELLOW = (255,255,0)
BLANK = (255, 255, 255)
BLUE = (0, 0, 255)
RED = (255, 0, 0)
DEPTH = 6

ROW_COUNT = 6
COLUMN_COUNT = 7

PLAYER = 0
AI = 1

EMPTY = 0
PLAYER_PIECE = 1
AI_PIECE = 2

WINDOW_LENGTH = 4

def create_board():
	board = np.zeros((ROW_COUNT,COLUMN_COUNT))
	return board

def drop_piece(board, row, col, piece):
	board[row][col] = piece

def is_valid_location(board, col):
	return board[ROW_COUNT-1][col] == 0

def get_next_open_row(board, col):
	for r in range(ROW_COUNT):
		if board[r][col] == 0:
			return r

def print_board(board):
	print(np.flip(board, 0))

def winning_move(board, piece):
	for c in range(COLUMN_COUNT-3):
		for r in range(ROW_COUNT):
			if board[r][c] == piece and board[r][c+1] == piece and board[r][c+2] == piece and board[r][c+3] == piece:
				return True

	for c in range(COLUMN_COUNT):
		for r in range(ROW_COUNT-3):
			if board[r][c] == piece and board[r+1][c] == piece and board[r+2][c] == piece and board[r+3][c] == piece:
				return True

	for c in range(COLUMN_COUNT-3):
		for r in range(ROW_COUNT-3):
			if board[r][c] == piece and board[r+1][c+1] == piece and board[r+2][c+2] == piece and board[r+3][c+3] == piece:
				return True

	for c in range(COLUMN_COUNT-3):
		for r in range(3, ROW_COUNT):
			if board[r][c] == piece and board[r-1][c+1] == piece and board[r-2][c+2] == piece and board[r-3][c+3] == piece:
				return True

def evaluate_window(window, piece):
	score = 0
	opp_piece = PLAYER_PIECE
	if piece == PLAYER_PIECE:
		opp_piece = AI_PIECE

	if window.count(piece) == 4:
		score += 100
	elif window.count(piece) == 3 and window.count(EMPTY) == 1:
		score += 5
	elif window.count(piece) == 2 and window.count(EMPTY) == 2:
		score += 2

	if window.count(opp_piece) == 3 and window.count(EMPTY) == 1:
		score -= 4

	return score

def score_position(board, piece):
	score = 0

	center_array = [int(i) for i in list(board[:, COLUMN_COUNT//2])]
	center_count = center_array.count(piece)
	score += center_count * 3

	for r in range(ROW_COUNT):
		row_array = [int(i) for i in list(board[r,:])]
		for c in range(COLUMN_COUNT-3):
			window = row_array[c:c+WINDOW_LENGTH]
			score += evaluate_window(window, piece)

	for c in range(COLUMN_COUNT):
		col_array = [int(i) for i in list(board[:,c])]
		for r in range(ROW_COUNT-3):
			window = col_array[r:r+WINDOW_LENGTH]
			score += evaluate_window(window, piece)

	for r in range(ROW_COUNT-3):
		for c in range(COLUMN_COUNT-3):
			window = [board[r+i][c+i] for i in range(WINDOW_LENGTH)]
			score += evaluate_window(window, piece)

	for r in range(ROW_COUNT-3):
		for c in range(COLUMN_COUNT-3):
			window = [board[r+3-i][c+i] for i in range(WINDOW_LENGTH)]
			score += evaluate_window(window, piece)

	return score

def is_terminal_node(board):
	return winning_move(board, PLAYER_PIECE) or winning_move(board, AI_PIECE) or len(get_valid_locations(board)) == 0

def minimax(board, depth, alpha, beta, maximizingPlayer):
	valid_locations = get_valid_locations(board)
	is_terminal = is_terminal_node(board)
	if depth == 0 or is_terminal:
		if is_terminal:
			if winning_move(board, AI_PIECE):
				return (None, 100000000000000)
			elif winning_move(board, PLAYER_PIECE):
				return (None, -10000000000000)
			else:
				return (None, 0)
		else:
			return (None, score_position(board, AI_PIECE))
	if maximizingPlayer:
		value = -math.inf
		column = random.choice(valid_locations)
		for col in valid_locations:
			row = get_next_open_row(board, col)
			b_copy = board.copy()
			drop_piece(b_copy, row, col, AI_PIECE)
			new_score = minimax(b_copy, depth-1, alpha, beta, False)[1]
			if new_score > value:
				value = new_score
				column = col
			alpha = max(alpha, value)
			if alpha >= beta:
				break
		return column, value

	else:
		value = math.inf
		column = random.choice(valid_locations)
		for col in valid_locations:
			row = get_next_open_row(board, col)
			b_copy = board.copy()
			drop_piece(b_copy, row, col, PLAYER_PIECE)
			new_score = minimax(b_copy, depth-1, alpha, beta, True)[1]
			if new_score < value:
				value = new_score
				column = col
			beta = min(beta, value)
			if alpha >= beta:
				break
		return column, value

def get_valid_locations(board):
	valid_locations = []
	for col in range(COLUMN_COUNT):
		if is_valid_location(board, col):
			valid_locations.append(col)
	return valid_locations

def pick_best_move(board, piece):

	valid_locations = get_valid_locations(board)
	best_score = -10000
	best_col = random.choice(valid_locations)
	for col in valid_locations:
		row = get_next_open_row(board, col)
		temp_board = board.copy()
		drop_piece(temp_board, row, col, piece)
		score = score_position(temp_board, piece)
		if score > best_score:
			best_score = score
			best_col = col

	return best_col

board = create_board()
game_over = False


turn = random.randint(PLAYER, AI)

def getTurn():
    try:
        if pyautogui.pixelMatchesColor(1027, 930, BLANK) and pyautogui.pixelMatchesColor(1019, 924, BLANK) and pyautogui.pixelMatchesColor(893, 924, BLANK):
            return True
    except:
        if pyautogui.pixelMatchesColor(1027, 930, BLANK) and pyautogui.pixelMatchesColor(1019, 924, BLANK) and pyautogui.pixelMatchesColor(893, 924, BLANK):
            return True
    return False

def isRed(x, y):
    try:
        return pyautogui.pixelMatchesColor(x, y, RED)
    except:
        return pyautogui.pixelMatchesColor(x, y, RED)

def isBlue(x, y):
    try:
        return pyautogui.pixelMatchesColor(x, y, BLUE)
    except:
        return pyautogui.pixelMatchesColor(x, y, BLUE)

def updateBoardState(boardState):
    for i in range(ROW_COUNT):
        for j in range(COLUMN_COUNT):
            if isRed(BOARDPOS[i][j][0], BOARDPOS[i][j][1]):
                boardState[ROW_COUNT - i - 1][j] = 2
            elif isBlue(BOARDPOS[i][j][0], BOARDPOS[i][j][1]):
                boardState[ROW_COUNT - i - 1][j] = 1

def mouseClick(coords):
    win32api.SetCursorPos((coords[0],coords[1]))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)

if getTurn():
	turn = PLAYER
else:
	turn = AI

while game_over == False:
	if getTurn():
		turn = AI
		updateBoardState(board)
	if turn == AI and not game_over:				
		col = random.randint(0, COLUMN_COUNT-1)
		col = pick_best_move(board, AI_PIECE)
		col, minimax_score = minimax(board, DEPTH, -math.inf, math.inf, True)

		if is_valid_location(board, col):
			row = get_next_open_row(board, col)
			mouseClick(BOARDPOS[row][col])

			if winning_move(board, AI_PIECE):
				game_over = True


			turn = PLAYER
			updateBoardState(board)
		if keyboard.is_pressed('q'):
			game_over = True
			break
	if keyboard.is_pressed('q'):
		game_over = True