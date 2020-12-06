import xlsxwriter
import json
from datetime import datetime


class Excella:
    """
        Renders schedules in a Excel spreadsheet
        Properties:
            1. Renders schedule 'blocks' per section
            3. Load officer data into scheds as metadata
    """

    SCHED_BLOCKS_PER_SHEET = 12
    HORIZONTAL_CELL_OFFSET = 3
    VERTICAL_CELL_OFFSET = 2

    # Column positions
    SCHED_BLOCK_COLUMN_START = 1
    SCHED_BLOCK_COLUMN_END = 7

    # Cell widths in pixels
    TIME_CELL_WIDTH = 12.57
    SCHED_CELL_WIDTH = 25.00

    # Day strings
    DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

    # Color hex strings
    COLORS = ["#70AD47", "#B4C6E7", "#ED7D31",
              "#F4B084", "#FFE699", "#C00000", "#4BACC6"]

    # Time intervals
    INTERVALS = {
        "07:00 AM": 0,
        "07:30 AM": 1,
        "08:00 AM": 2,
        "08:30 AM": 3,
        "09:00 AM": 4,
        "09:30 AM": 5,
        "10:00 AM": 6,
        "10:30 AM": 7,
        "11:00 AM": 8,
        "11:30 AM": 9,
        "12:00 PM": 10,
        "12:30 PM": 11,
        "01:00 PM": 12,
        "01:30 PM": 13,
        "02:00 PM": 14,
        "02:30 PM": 15,
        "03:00 PM": 16,
        "03:30 PM": 17,
        "04:00 PM": 18,
        "04:30 PM": 19,
        "05:00 PM": 20,
        "05:30 PM": 21,
        "06:00 PM": 22
    }

    def __init__(self, entries, officers_path, export_name):
        self.officers_path = officers_path
        self.workbook = xlsxwriter.Workbook(export_name)
        self.entries = entries
        self.has_set_columns = False

    def begin(self):
        """
            Begins drawing cycle
        """
        self.__log("Beginning drawing cycle...")

        # Set column stylings
        worksheet = self.workbook.add_worksheet()
        self.__set_columns(worksheet)

        # Begin drawing
        offset = Excella.HORIZONTAL_CELL_OFFSET

        # Open officers JSON file
        f = open(self.officers_path)
        officers_json = json.load(f)
        idx = 0

        for entry in officers_json:
            section = entry["section"]
            officers = entry["officers"]

            # Draw schedule of section and its officers
            offset = self.__draw(worksheet, offset, section, officers) + 2
            self.__log(f"Drew officers schedule of {section}")

        # Close file stream
        f.close()

    def __draw(self, worksheet, offset, section, officers):
        """
            Draws a schedule block given a section and its officers
        """
        buckets = self.__load_buckets(section)
        curr_offset = offset

        # Draw metadata
        curr_offset = self.__draw_metadata(
            worksheet, curr_offset, section, officers)

        # Draw railings
        next_offset = self.__draw_railings(worksheet, curr_offset)

        # Draw schedule blocks
        self.__draw_schedules(worksheet, curr_offset, buckets)

        return next_offset + 1

    def __draw_metadata(self, worksheet, offset, section, officers):
        """
            Draws schedule metadata
        """
        curr_row = offset
        curr_col = Excella.SCHED_BLOCK_COLUMN_START

        # Draw section text
        section_format = self.workbook.add_format()
        section_format.set_bold(True)
        section_format.set_align("center")

        worksheet.write(curr_row, curr_col, section, section_format)

        # Draw position text
        pos_col = Excella.SCHED_BLOCK_COLUMN_START + 1
        officers_col = Excella.SCHED_BLOCK_COLUMN_START + 3

        pos_format = self.workbook.add_format()
        pos_format.set_bold(True)

        for officer in officers:
            position = officer["position"]
            names = officer["names"]

            worksheet.write(curr_row, pos_col, position, pos_format)
            worksheet.write(curr_row, officers_col, ", ".join(names))
            curr_row += 1

        return curr_row + 1

    def __draw_schedules(self, worksheet, offset, buckets):
        """
            Draws the schedule blocks for the given bucket
        """
        curr_row = offset + 1
        colors = {}
        color_index = 0

        for i in range(len(buckets)):
            curr_col = Excella.SCHED_BLOCK_COLUMN_START + 1 + i
            day = buckets[i]

            for sch_obj in day:

                # Get new color if not yet registered
                code = sch_obj["code"]
                if code not in colors:
                    color = Excella.COLORS[color_index]
                    colors[code] = color
                    color_index += 1
                else:
                    color = colors[code]

                # Draw schedule cell
                self.__draw_sched_cell(
                    worksheet, curr_row, curr_col, sch_obj, color)

    def __draw_sched_cell(self, worksheet, row, col, sch_obj, color):
        """
            Draws a schedule cell with its subject code and time
        """
        sched_format = self.workbook.add_format()
        sched_format.set_bg_color(color)

        # Draw top part of the cell
        start = sch_obj["time_start_interval"]
        end = sch_obj["time_end_interval"]

        curr_row = row + start
        worksheet.write(curr_row, col, "", sched_format)

        # Draw cell name
        curr_row += 1
        worksheet.write(curr_row, col, sch_obj["code"], sched_format)

        # Draw cell time
        curr_row += 1
        worksheet.write(curr_row, col, sch_obj["time"], sched_format)

        curr_row += 1
        for i in range(curr_row, curr_row + (end - start - 2)):
            worksheet.write(i, col, "", sched_format)

    def __draw_railings(self, worksheet, offset):
        """
            Draws the time rows and day columns
        """
        workbook = self.workbook
        curr_row = offset + 1
        next_row = curr_row

        # Draw time rows... 7:30 am to 5:00 pm rows
        time_format = workbook.add_format()
        time_format.set_align("right")
        time_format.set_indent(1)

        for time in Excella.INTERVALS.keys():
            worksheet.write(
                curr_row, Excella.SCHED_BLOCK_COLUMN_START, time, time_format)
            curr_row += 1
            next_row = curr_row

        # Draw day columns... Monday to Saturday
        curr_row = offset
        curr_col = Excella.SCHED_BLOCK_COLUMN_START + 1
        day_format = workbook.add_format()
        day_format.set_bold(True)
        day_format.set_align("center")

        for day in Excella.DAYS:
            worksheet.write(curr_row, curr_col, day, day_format)
            curr_col += 1

        return next_row

    def __load_buckets(self, section_name):
        """
            Loads the schedule objects into their corresponding array buckets. Returns the resulting bucket
        """
        section = self.entries[section_name]
        buckets = [[] for _ in range(len(Excella.DAYS))]

        # Access subjects
        for subject in section:

            # Access schedules
            schedules = subject["schedules"]

            for day in schedules:
                sch = schedules[day]

                # Create schedule object
                sch_obj = {}
                sch_obj["name"] = subject["name"]
                sch_obj["code"] = subject["code"]
                sch_obj["room"] = sch["room"]

                start = self.__convert_time(sch["time_start"])
                end = self.__convert_time(sch["time_end"])

                sch_obj["time"] = f"{start} - {end}"
                sch_obj["time_start_interval"] = self.__get_interval(start)
                sch_obj["time_end_interval"] = self.__get_interval(end)

                # Insert into day bucket
                bucket = buckets[Excella.DAYS.index(day)]
                bucket.append(sch_obj)

                # Sort by time start
                bucket.sort(
                    key=lambda x: x["time_start_interval"], reverse=False)

        return buckets

    def __set_columns(self, worksheet):
        """
            Set the column styles
        """
        if not self.has_set_columns:
            # Set column widths
            start = Excella.SCHED_BLOCK_COLUMN_START
            worksheet.set_column(start, start, Excella.TIME_CELL_WIDTH)
            worksheet.set_column(
                start + 1, Excella.SCHED_BLOCK_COLUMN_END, Excella.SCHED_CELL_WIDTH)

            self.__log("Column stylings has been set")
            self.has_set_columns = True

    def __get_interval(self, time_str):
        """
            Gets the integer interval of a time string. Prone to error correction if time string is out of bounds
        """
        try:
            interval = Excella.INTERVALS[time_str]
        except KeyError:
            # Get nearest interval if not explicitly found in dictionary
            dt = self.__to_dt(time_str)

            # Switch AM to PM if its earlier than 7:00 AM
            if dt < self.__to_dt("07:00 AM"):
                time_str = time_str.replace("AM", "PM").strip()
                interval = Excella.INTERVALS[time_str]
            else:
                for k in Excella.INTERVALS.keys():
                    if self.__to_dt(k) >= dt:
                        interval = Excella.INTERVALS[k]

        return interval

    def __convert_time(self, military_time):
        """
            Converts military time to standard time string
        """
        return datetime.strptime(military_time, '%H:%M:%S').strftime('%I:%M %p').strip()

    def __to_dt(self, time):
        """
            Converts time string to standard date time object
        """
        return datetime.strptime(time, '%I:%M %p')

    def __log(self, msg):
        """
            Simple logging
        """
        print(f"[3xc311a]: {msg}")

    def close(self):
        """
            Closes the workbook file stream
        """
        self.workbook.close()
