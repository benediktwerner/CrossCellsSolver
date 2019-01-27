import cv2
import numpy as np
import math
import mss
import pytesseract
import time
import win32api
import win32con
import win32gui
from z3 import Bool, If, simplify, sat, Solver


BORDER_TOP = 100
BORDER_LEFT = 100
BORDER_BOTTOM = 100
BORDER_RIGHT = 100
CELL_SIZE = 35
NUMBER_ONE_WIDTH = 11

WINDOW_HANDLE = win32gui.FindWindow(None, "CrossCells")
LEFT, TOP, RIGHT, BOTTOM = win32gui.GetWindowRect(WINDOW_HANDLE)
MONITOR = {
    "left": LEFT + BORDER_LEFT,
    "top": TOP + BORDER_TOP,
    "width": RIGHT - LEFT - BORDER_LEFT - BORDER_RIGHT,
    "height": BOTTOM - TOP - BORDER_TOP - BORDER_BOTTOM
}


class Rect:
    def __init__(self, rect):
        self.x = rect[0]
        self.y = rect[1]
        self.width = rect[2]
        self.height = rect[3]
        self.x2 = self.x + self.width
        self.y2 = self.y + self.height
        self.center_x = self.x + self.width // 2
        self.center_y = self.y + self.height // 2

    def to_rect(self, border=0):
        return (self.x-border, self.y-border, self.width+border*2, self.height+border*2)

    def dist(self, other):
        return math.sqrt((self.center_x - other.center_x)**2 + (self.center_y - other.center_y)**2)


class Label(Rect):
    def __init__(self, rect):
        super().__init__(rect)
        self.connected = []
        self.merged = False
        self.text = None

    def __merge(self, other):
        self.x = min(self.x, other.x)
        self.y = min(self.y, other.y)
        self.x2 = max(self.x2, other.x2)
        self.y2 = max(self.y2, other.y2)

    def merge(self):
        self.merged = True
        for other in self.connected:
            if other.merged:
                continue
            other.merge()
            self.__merge(other)

        self.width = self.x2 - self.x
        self.height = self.y2 - self.y
        self.center_x = self.x + self.width // 2
        self.center_y = self.y + self.height // 2


class Cell(Rect):
    def __init__(self, rect, text):
        super().__init__(rect)
        self.text = text
        self.square = None
        self.variable = Bool("{},{}".format(self.x, self.y))


class Square(Rect):
    def __init__(self, rect):
        super().__init__(rect)
        self.cells = []
        self.text = None

    def get_constraint(self):
        if "[" in self.text:
            return sum(If(cell.variable, 1, 0) for cell in self.cells) == int(self.text.strip("[]"))

        term = 0
        for cell in self.cells:
            if cell.text[0] == "+":
                term += If(cell.variable, int(cell.text[1:]), 0)
            else:
                term *= If(cell.variable, int(cell.text[1:]), 1)
        return term == int(self.text)


class Line:
    def __init__(self, rect, cell=None):
        self.x = rect[0] + rect[2] // 2
        self.y = rect[1] + rect[3] // 2
        self.center_x = self.x
        self.center_y = self.y
        self.cells = [] if cell is None else [cell]
        self.horizontal = None
        self.text = None

    def order_cells(self):
        if max(abs(cell.center_x - self.x) for cell in self.cells) < CELL_SIZE:
            self.cells.sort(key=lambda c: c.y, reverse=self.cells[0].y > self.y)
            self.horizontal = False
        else:
            self.cells.sort(key=lambda c: c.x, reverse=self.cells[0].x > self.x)
            self.horizontal = True

    def dist(self, other):
        return math.sqrt((self.center_x - other.center_x)**2 + (self.center_y - other.center_y)**2)

    def get_constraint(self):
        if "[" in self.text:
            return sum(If(cell.variable, 1, 0) for cell in self.cells) == int(self.text.strip("[]"))

        term = 0
        for cell in self.cells:
            if cell.text[0] == "+":
                term += If(cell.variable, int(cell.text[1:]), 0)
            else:
                term *= If(cell.variable, int(cell.text[1:]), 1)
        return term == int(self.text)


def detect_text(img, rect):
    x, y, w, h = rect
    return pytesseract.image_to_string(img[y:y+h, x:x+w], config="tesseract.conf")


def findBoundingRects(img, inside=False):
    contours, _ = cv2.findContours(img, cv2.RETR_LIST if inside else cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    return [cv2.boundingRect(contour) for contour in contours]


def screenshot():
    with mss.mss() as sct:
        return np.array(sct.grab(MONITOR))


def solve():
    img = screenshot()

    img_orig = img.copy()
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, img = cv2.threshold(img, 110, 255, cv2.THRESH_BINARY_INV)

    cells = []
    squares = []
    lines = []
    labels = []

    for rect in findBoundingRects(img):
        x, y, w, h = rect

        img_orig = cv2.rectangle(img_orig, rect, (0, 0, 255))
        if w > CELL_SIZE and h > CELL_SIZE:
            text = detect_text(img, rect)
            img_orig = cv2.putText(img_orig, text, (x, y), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
            cells.append(Cell(rect, text))
        else:
            label = Label(rect)
            labels.append(label)
            for other in labels:
                if label.dist(other) < 30:
                    label.connected.append(other)
                    other.connected.append(label)

    labels_merged = []
    for label in labels:
        if label.merged:
            continue
        label.merge()
        labels_merged.append(label)
        label.text = detect_text(img, label.to_rect(10))
    labels = labels_merged
    all_labels = labels[:]

    for cell in cells:
        win32api.SetCursorPos((cell.center_x + LEFT + BORDER_LEFT, cell.center_y + TOP + BORDER_TOP))
        time.sleep(0.5)

        img2 = screenshot()
        img2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        _, img2 = cv2.threshold(img2, 120, 255, cv2.THRESH_BINARY_INV)
        diff = img2-img

        for cell2 in cells:
            diff[cell2.y-5:cell2.y2+5, cell2.x-5:cell2.x2+5] = 0
        for label in all_labels:
            diff[label.y-3:label.y2+3, label.x-3:label.x2+3] = 0

        _, diff = cv2.threshold(diff, 110, 255, cv2.THRESH_BINARY)

        for rect in findBoundingRects(diff, inside=True):
            if rect[2] > CELL_SIZE and rect[3] > CELL_SIZE:
                if cell.square is None:
                    square = Square(rect)
                    for cell2 in cells:
                        if square.x < cell2.x < square.x2 and square.y < cell2.y < square.y2:
                            square.cells.append(cell2)
                            cell2.square = square
                    squares.append(square)
            elif abs(cell.center_x - rect[0] - rect[2] // 2) < 5 or abs(cell.center_y - rect[1] - rect[3] // 2) < 5:
                for line in lines:
                    if abs(line.x - rect[0]) < 5 and abs(line.y - rect[1]) < 5:
                        line.cells.append(cell)
                        break
                else:
                    line = Line(rect, cell)
                    label = min(labels, key=line.dist)
                    labels.remove(label)
                    line.text = label.text
                    lines.append(line)
                    img_orig = cv2.line(img_orig, (line.center_x, line.center_y),
                                        (label.center_x, label.center_y), (255, 0, 0))
                    img_orig = cv2.rectangle(img_orig, label.to_rect(), (255, 0, 0))

    for square in squares:
        for label in labels:
            if square.x < label.x < square.x2 and square.y < label.y < square.y2:
                square.text = label.text
                labels.remove(label)
                img_orig = cv2.line(img_orig, (square.x, square.y), (label.center_x, label.center_y), (0, 255, 0))
                img_orig = cv2.rectangle(img_orig, label.to_rect(), (0, 255, 0))
                break
        img_orig = cv2.putText(img_orig, square.text, (square.x, square.y), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)

    if labels:
        for label in labels:
            img_orig = cv2.rectangle(img_orig, label.to_rect(), (0, 0, 255))

    for line in lines:
        line.order_cells()
        x = line.x
        y = line.y
        img_orig = cv2.putText(img_orig, line.text, (x, y), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
        for cell in line.cells:
            img_orig = cv2.line(img_orig, (x, y), (cell.center_x, cell.center_y), (0, 0, 255))
            x = cell.center_x
            y = cell.center_y

    for cell in cells:
        if cell.square is not None:
            img_orig = cv2.line(img_orig, (cell.x, cell.y), (cell.square.x, cell.square.y), (0, 255, 0))

    solver = Solver()

    for line in lines:
        solver.add(simplify(line.get_constraint()))

    for square in squares:
        solver.add(simplify(square.get_constraint()))

    if solver.check() == sat:
        model = solver.model()
        for cell in cells:
            x = cell.center_x + LEFT + BORDER_LEFT
            y = cell.center_y + TOP + BORDER_TOP
            win32api.SetCursorPos((x, y))
            time.sleep(0.01)
            if model[cell.variable]:
                win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
                time.sleep(0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, x, y, 0, 0)
            else:
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
                time.sleep(0.01)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
            time.sleep(0.01)
        return True

    print(solver)
    cv2.imshow("Detection", img_orig)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    return False


time.sleep(5)

for _ in range(5):
    solve()
    win32api.SetCursorPos((TOP, LEFT))
    time.sleep(7)
