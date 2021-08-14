import os
import xlrd
import datetime as dt
import keyboard

# How to đặt tên biến :(

WEEKDAY_DICT = {
    'Thứ 2': 0,
    'Thứ 3': 1,
    'Thứ 4': 2,
    'Thứ 5': 3,
    'Thứ 6': 4,
    'Thứ 7': 5,
    'Chủ nhật': 6,
}


class Timetable:
    def __init__(self, path: str):
        self.exel = xlrd.open_workbook(path)
        self.sheet = self.exel.sheet_by_index(0)

        self.first_row = 10
        self.last_row = self.sheet.nrows - 9
        self.subject_col = 5
        self.time_col = 7

    @staticmethod
    def process_data(time_loc_data: str) -> list:
        # Giả sử ta có đầu vào time_loc_data như sau:

        # Từ 16/08/2021 đến 29/08/2021:
        #  Thứ 5 tiết 6,7 tại C3.201 ID 270 224 0011 C3
        #  Thứ 6 tiết 1,2 tại C3.201 ID 270 224 0011 C3
        # Từ 06/09/2021 đến 31/10/2021:
        #  Thứ 5 tiết 6,7 tại C3.201 ID 270 224 0011 C3
        #  Thứ 6 tiết 1,2 tại C3.201 ID 270 224 0011 C3

        # Ta cần cắt thành 2 cụm bởi string 'Từ '
        # Mỗi cụm này sẽ được lưu ở 1 dictionary có dạng:
        # data_dict = {
        #   'date': '16/08/2021 đến 29/08/2021'
        #   'time': ['Thứ 5 tiết 6,7 tại C3.201 ID 270 224 0011 C3', 'Thứ 6 tiết 1,2 tại C3.201 ID 270 224 0011 C3']
        # }

        # Từ các dict này ta kết hợp thành 1 list và return ra dữ liệu đã được xử lý

        list_temp = time_loc_data.split('Từ ')[1:]

        result = []
        for element in list_temp:
            temp2 = element.split('\n')

            data_dict = {
                'date': temp2[0][:-1]
            }

            weekday_list = []
            for i in range(1, len(temp2)):
                if temp2[i] != '':
                    weekday_list.append(temp2[i][1:])

            data_dict['time'] = weekday_list

            result.append(data_dict)

        return result

    def get_current_subjects(self, current_date: dt):
        free = True
        for cur_row in range(self.first_row, self.last_row + 1):
            subject_cell = self.sheet.cell(rowx=cur_row, colx=self.subject_col)
            time_and_location_cell = self.sheet.cell(rowx=cur_row, colx=self.time_col)

            processed_data = self.process_data(time_and_location_cell.value)

            for data in processed_data:
                time_location_list = self.check(data, current_date)

                if len(time_location_list) > 0:
                    print(f'\n{subject_cell.value}')
                    for lc in time_location_list:
                        print(f'\t+ {lc}')
                    free = False

        if free is True:
            print('\nHôm nay bạn rảnh :)')

    @staticmethod
    def get_date_from_string(date: str) -> dt:
        temp = date.split('/')
        return dt.datetime(year=int(temp[2]), month=int(temp[1]), day=int(temp[0]))

    def check(self, time_data: dict, current_date: dt) -> list:
        subject_date = time_data['date']

        start_date = self.get_date_from_string(subject_date[0: 10])
        end_date = self.get_date_from_string(subject_date[15: 25])

        result = []
        for time in time_data['time']:
            subject_weekday = WEEKDAY_DICT[time[:5]]

            if start_date <= current_date <= end_date and subject_weekday == current_date.weekday():
                result.append(time)

        return result


today = dt.datetime.now()

try:
    timetable = Timetable('ThoiKhoaBieuSinhVien.xls')
except FileNotFoundError:
    print('ERROR - Không tìm thấy file ThoiKhoaBieuSinhVien.xls')
    print('      - Hãy tải file tkb trên trang dangkytinchi và bỏ vào cùng chỗ với file timetable.exe')
    input()
    exit()


def print_timetable():
    os.system('cls')
    print(f'\t\t\t{today.strftime("%A - %d/%m/%Y")}')
    timetable.get_current_subjects(current_date=today)


def handle_left_key():
    global today
    today = today - dt.timedelta(days=1)
    print_timetable()


def handle_right_key():
    global today
    today = today + dt.timedelta(days=1)
    print_timetable()


if __name__ == '__main__':
    print_timetable()

    keyboard.add_hotkey('left', handle_left_key)
    keyboard.add_hotkey('right', handle_right_key)

    keyboard.wait()
