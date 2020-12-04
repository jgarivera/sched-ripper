import xlsxwriter

class Excella:
    """
        Renders schedules in a Excel spreadsheet
        Properties:
            1. Renders schedule 'blocks' per section
            2. Limit number of schedule blocks per sheet
            3. Load officer data into scheds
    """

    SCHED_BLOCKS_PER_SHEET = 12
    HORIZONTAL_CELL_OFFSET = 3
    VERTICAL_CELL_OFFSET = 2

    def __init__(self, entries, export_name):
        self.workbook = xlsxwriter.Workbook(export_name)
        self.entries = entries
        self.has_set_columns = False
        self.sheet_count = 0
        self.sheet_sched_blocks_count = 0
        self.total_sched_blocks_count = 0

    def draw(self, section, metadata):
        """
            Draws a sched block
        """
        worksheet = self.workbook.add_worksheet()

        if not self.has_set_columns:
            # Set column widths
            worksheet.set_cols
        for entry in self.entries:
            


    def close(self):
        self.workbook.close()