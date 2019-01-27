import time
import math

import cv2
import numpy as np
import mss
import pytesseract
import win32con
import win32api
import win32gui
from z3 import Bool, If, simplify, sat, Solver


BORDER_TOP = 100
BORDER_LEFT = 100
BORDER_BOTTOM = 100
BORDER_RIGHT = 100
CELL_SIZE = 35


class Rect:
    def __init__(self, rect):
        if isinstance(rect, Rect):
            self.x = rect.x
            self.y = rect.y
            self.width = rect.width
            self.height = rect.height
        else:
            self.x = rect[0]
            self.y = rect[1]
            self.width = rect[2]
            self.height = rect[3]
        self.x2 = self.x + self.width
        self.y2 = self.y + self.height
        self.center_x = self.x + self.width // 2
        self.center_y = self.y + self.height // 2

    @staticmethod
    def from_corner_rect(rect):
        x1, y1, x2, y2 = rect
        return Rect((x1, y1, x2-x1, y2-y1))

    def enlarge(self, border):
        return Rect((self.x-border, self.y-border, self.width+border*2, self.height+border*2))

    def to_rect(self):
        return (self.x, self.y, self.width, self.height)

    def to_slice(self):
        return (slice(self.y, self.y2), slice(self.x, self.x2))

    def dist(self, other):
        return math.sqrt((self.center_x - other.center_x)**2 + (self.center_y - other.center_y)**2)

    def contains(self, other):
        return self.x < other.x and other.x2 < self.x2 and self.y < other.y and other.y2 < self.y2


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


class Constraint(Rect):
    def __init__(self, rect, initial_cells=None):
        super().__init__(rect)
        self.cells = [] if initial_cells is None else initial_cells
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


class LineConstraint(Constraint):
    def __init__(self, rect, cell=None):
        super().__init__(rect, None if cell is None else [cell])

    def sort_cells(self):
        if max(abs(cell.center_x - self.x) for cell in self.cells) < CELL_SIZE:
            self.cells.sort(key=lambda c: c.y, reverse=self.cells[0].y > self.y)
        else:
            self.cells.sort(key=lambda c: c.x, reverse=self.cells[0].x > self.x)


def find_bounding_rects(img, inside=False):
    mode = cv2.RETR_LIST if inside else cv2.RETR_EXTERNAL
    contours, _ = cv2.findContours(img, mode, cv2.CHAIN_APPROX_SIMPLE)
    return [Rect(cv2.boundingRect(contour)) for contour in contours]


class CrossCellsSolver:
    def __init__(self):
        self.window_handle = win32gui.FindWindow(None, "CrossCells")
        self.window_rect = Rect.from_corner_rect(win32gui.GetWindowRect(self.window_handle))
        self.monitor = {
            "left": self.window_rect.x + BORDER_LEFT,
            "top": self.window_rect.y + BORDER_TOP,
            "width": self.window_rect.width - BORDER_LEFT - BORDER_RIGHT,
            "height": self.window_rect.height - BORDER_TOP - BORDER_BOTTOM
        }
        self.img = None
        self.img_orig = None
        self.cells = []
        self.squares = []
        self.lines = []
        self.labels = []

    def move_mouse(self, x=0, y=0):
        win32api.SetCursorPos((
            x + self.window_rect.x + BORDER_LEFT,
            y + self.window_rect.y + BORDER_TOP
        ))

    def click(self, x, y, right=False):
        self.move_mouse(x, y)
        time.sleep(0.01)

        x = x + self.window_rect.x + BORDER_LEFT
        y = y + self.window_rect.y + BORDER_TOP

        btn_down = win32con.MOUSEEVENTF_RIGHTDOWN if right else win32con.MOUSEEVENTF_LEFTDOWN
        btn_up = win32con.MOUSEEVENTF_RIGHTUP if right else win32con.MOUSEEVENTF_LEFTUP

        win32api.mouse_event(btn_down, x, y, 0, 0)
        time.sleep(0.01)
        win32api.mouse_event(btn_up, x, y, 0, 0)
        time.sleep(0.01)

    def screenshot(self):
        with mss.mss() as sct:
            return np.array(sct.grab(self.monitor))

    def detect_text(self, rect):
        return pytesseract.image_to_string(self.img[rect.to_slice()], config="tesseract.conf")

    def draw_line(self, start, end, color=(0, 0, 255)):
        self.img_orig = cv2.line(self.img_orig, start, end, color)

    def draw_rect(self, rect, color=(0, 0, 255)):
        self.img_orig = cv2.rectangle(self.img_orig, rect.to_rect(), color)

    def draw_text(self, text, rect, color=(0, 0, 255)):
        self.img_orig = cv2.putText(self.img_orig, text, (rect.x, rect.y), cv2.FONT_HERSHEY_SIMPLEX, 1, color, 2)

    def do_level(self):
        self.detect_level()
        self.solve_level()

    def detect_level(self):
        self.img = self.screenshot()
        self.img_orig = self.img.copy()
        self.img = cv2.cvtColor(self.img, cv2.COLOR_BGR2GRAY)
        _, self.img = cv2.threshold(self.img, 110, 255, cv2.THRESH_BINARY_INV)

        self.cells = []
        self.squares = []
        self.lines = []
        self.labels = []

        self.detect_objects()
        self.process_labels()
        self.process_cells()
        self.process_lines()
        self.process_squares()

        if self.labels:
            for label in self.labels:
                self.draw_rect(label, (0, 0, 255))

    def detect_objects(self):
        for rect in find_bounding_rects(self.img):
            if rect.width > CELL_SIZE and rect.height > CELL_SIZE:
                text = self.detect_text(rect)
                self.draw_text(text, rect)
                self.cells.append(Cell(rect, text))
            else:
                label = Label(rect)
                self.labels.append(label)
                for other in self.labels:
                    if label.dist(other) < 30:
                        label.connected.append(other)
                        other.connected.append(label)

    def process_labels(self):
        labels_merged = []
        for label in self.labels:
            if label.merged:
                continue
            label.merge()
            labels_merged.append(label)
            label.text = self.detect_text(label.enlarge(10))
        self.labels = labels_merged

    def process_cells(self):
        for cell in self.cells:
            self.move_mouse(cell.center_x, cell.center_y)
            time.sleep(0.5)

            img2 = self.screenshot()
            img2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
            _, img2 = cv2.threshold(img2, 120, 255, cv2.THRESH_BINARY_INV)
            diff = img2 - self.img

            for cell2 in self.cells:
                diff[cell2.enlarge(5).to_slice()] = 0
            for label in self.labels:
                diff[label.enlarge(3).to_slice()] = 0

            _, diff = cv2.threshold(diff, 110, 255, cv2.THRESH_BINARY)

            for rect in find_bounding_rects(diff, inside=True):
                if rect.width > CELL_SIZE and rect.height > CELL_SIZE:
                    if cell.square is None:
                        self.add_square_constraint(rect)
                elif abs(cell.center_x - rect.center_x) < 5 or abs(cell.center_y - rect.center_y) < 5:
                    self.add_line_constraint(rect, cell)

    def add_square_constraint(self, rect):
        square = Constraint(rect)
        for cell in self.cells:
            if square.contains(cell):
                square.cells.append(cell)
                cell.square = square
        self.squares.append(square)

    def add_line_constraint(self, rect, cell):
        for line in self.lines:
            if rect.dist(line) < 5:
                line.cells.append(cell)
                break
        else:
            line = LineConstraint(rect, cell)
            self.lines.append(line)

    def process_lines(self):
        for line in self.lines:
            line.sort_cells()
            label = min(self.labels, key=line.dist)
            line.text = label.text
            self.labels.remove(label)

            self.draw_line((line.center_x, line.center_y), (label.center_x, label.center_y), (255, 0, 0))
            self.draw_rect(label, (255, 0, 0))

            x = line.x
            y = line.y
            self.draw_text(line.text, line, (0, 0, 255))
            for cell in line.cells:
                self.draw_line((x, y), (cell.center_x, cell.center_y), (0, 0, 255))
                x = cell.center_x
                y = cell.center_y

    def process_squares(self):
        for square in self.squares:
            for label in self.labels:
                if square.contains(label):
                    square.text = label.text
                    self.labels.remove(label)
                    self.draw_line((square.x, square.y), (label.center_x, label.center_y), (0, 255, 0))
                    self.draw_rect(label, (0, 255, 0))
                    break
            self.draw_text(square.text, square, (0, 255, 0))
            for cell in square.cells:
                self.draw_line((cell.x, cell.y), (cell.square.x, cell.square.y), (0, 255, 0))

    def solve_level(self):
        solver = Solver()

        for line in self.lines:
            solver.add(simplify(line.get_constraint()))

        for square in self.squares:
            solver.add(simplify(square.get_constraint()))

        if solver.check() == sat:
            model = solver.model()
            for cell in self.cells:
                self.click(cell.center_x, cell.center_y, model[cell.variable])
        else:
            print("Solver failed")
            print(solver)
            cv2.imshow("Detection", self.img_orig)
            cv2.waitKey(0)
            cv2.destroyAllWindows()


def main():
    time.sleep(5)
    solver = CrossCellsSolver()

    for _ in range(5):
        solver.do_level()
        solver.move_mouse()
        time.sleep(7)


if __name__ == "__main__":
    main()
