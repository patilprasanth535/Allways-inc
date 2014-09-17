from xlwt import Alignment, Borders, Font, Column


class ReportStyles():
    alignment = Alignment()
    font = Font()
    borders = Borders()
    # col = Column()

    def borders_light(self):
        self.borders.left = Borders.THIN
        self.borders.right = Borders.THIN
        self.borders.top = Borders.THIN
        self.borders.bottom = Borders.THIN
        return self.borders

    def align_hor_right(self):
        self.alignment.horz = Alignment.HORZ_RIGHT
        return self.alignment

    def align_hor_left(self):
        self.alignment.horz = Alignment.HORZ_LEFT
        return self.alignment

    def align_hor_center(self):
        self.alignment.horz = Alignment.HORZ_CENTER
        return self.alignment

    def text_bold(self):
        self.font.bold = True
        return self.font
    #
    # def column_width(self):
    #     self.col.set_width = 15
    #     return self.col