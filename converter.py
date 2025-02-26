import re
import tkinter
import json
from tkinter import *

import pyperclip


class Application:
    def __init__(self):
        self._root = Tk()
        self._root.resizable(False, False)
        self._root.title("Excel-Python Coverter")
        self.root = Frame(master=self._root, width=1200, height=600)
        self.root.pack()
        self.__layout()

    def __layout(self):
        self.info = Label(self.root, text="Author: Rebooting\nLast Update: 2025/2/26\nCurrent Version: 1.2 ",
                          relief=tkinter.SUNKEN)
        self.info.place(x=510, y=10, width=180, height=50)
        self.input_area = Text(self.root, bg="seashell")
        self.input_area.place(x=0, y=0, width=500, height=600)
        self.output_area = Text(self.root, bg="azure")
        self.output_area.place(x=700, y=0, width=500, height=600)
        self.label_1 = Button(self.root, text="Excel -----> Python", command=self.excel2python)
        self.label_1.place(x=520, y=100, width=160, height=30)

        self.output_format = IntVar()
        self.type_list = Radiobutton(self.root, text="list", value=1, variable=self.output_format)
        self.type_list.place(x=520, y=150)
        self.type_tuple = Radiobutton(self.root, text="tuple", value=2, variable=self.output_format, state="disabled")
        self.type_tuple.place(x=580, y=150)
        self.type_dict = Radiobutton(self.root, text="dict", value=3, variable=self.output_format, state="disabled")
        self.type_dict.place(x=640, y=150)
        self.output_format.set(1)

        self.with_header = Label(self.root, text="with header")
        self.with_header.place(x=520, y=182)
        self.output_with_names_or_not = IntVar()
        self.with_header_yes = Radiobutton(self.root, text="yes", value=1, variable=self.output_with_names_or_not)
        self.with_header_yes.place(x=600, y=180)
        self.with_header_no = Radiobutton(self.root, text="no", value=2, variable=self.output_with_names_or_not)
        self.with_header_no.place(x=650, y=180)
        self.output_with_names_or_not.set(2)
        self.output_default_prefix = 'Output_'
        self.output_counter = 0

        self.direction = Label(self.root, text="direction")
        self.direction.place(x=520, y=214)
        self.data_direction = IntVar()
        self.direction_vertical = Radiobutton(self.root, text="vertical", value=1, variable=self.data_direction)
        self.direction_vertical.place(x=600, y=214)
        self.direction_horizontal = Radiobutton(self.root, text="horizontal", value=2, variable=self.data_direction)
        self.direction_horizontal.place(x=600, y=235)
        self.data_direction.set(1)

        self.label_2 = Button(self.root, text="Excel  <-----  Python", command=self.python2excel)
        self.label_2.place(x=520, y=300, width=160, height=30)

        self.auto_copy = Label(self.root, text="auto copy")
        self.auto_copy.place(x=520, y=400)
        self.auto_copy_or_not = IntVar()
        self.auto_copy_yes = Radiobutton(self.root, text="yes", value=1, variable=self.auto_copy_or_not)
        self.auto_copy_yes.place(x=600, y=400)
        self.auto_copy_no = Radiobutton(self.root, text="no", value=2, variable=self.auto_copy_or_not)
        self.auto_copy_no.place(x=650, y=400)
        self.auto_copy_or_not.set(1)

        self.auto_copy = Label(self.root, text="number->str")
        self.auto_copy.place(x=510, y=425)
        self.number_convert_to_string_or_not = IntVar()
        self.auto_copy_yes = Radiobutton(self.root, text="yes", value=1, variable=self.number_convert_to_string_or_not)
        self.auto_copy_yes.place(x=600, y=425)
        self.auto_copy_no = Radiobutton(self.root, text="no", value=0, variable=self.number_convert_to_string_or_not)
        self.auto_copy_no.place(x=650, y=425)
        self.number_convert_to_string_or_not.set(0)

        self.auto_copy = Label(self.root, text="hierarchical")
        self.auto_copy.place(x=515, y=450)
        self.has_hierarchical_structure_or_not = IntVar()
        self.has_hierarchical_structure_or_not.trace_add('write', self.on_radiobutton_change)
        self.auto_copy_yes = Radiobutton(self.root, text="yes", value=1, variable=self.has_hierarchical_structure_or_not)
        self.auto_copy_yes.place(x=600, y=450)
        self.auto_copy_no = Radiobutton(self.root, text="no", value=0, variable=self.has_hierarchical_structure_or_not)
        self.auto_copy_no.place(x=650, y=450)
        self.has_hierarchical_structure_or_not.set(0)
        pass

    def on_radiobutton_change(self, *args):
        selected_value =self.has_hierarchical_structure_or_not.get()
        if selected_value:
            self.data_direction.set(1)
            self.direction_vertical.config(state="disabled")
            self.direction_horizontal.config(state="disabled")
        else:
            self.direction_vertical.config(state="normal")
            self.direction_horizontal.config(state="normal")
        pass

    @staticmethod
    def __find_all_substring_indices(text, substring):
        return [match.start() for match in re.finditer(re.escape(substring), text)]

    @staticmethod
    def __data_parser(_raw_data):
        ret_list = list()
        if _raw_data:
            if _raw_data[-1] == '\n':
                _raw_data = _raw_data[:-1]
            if _raw_data[-1] == '\n':
                _raw_data = _raw_data[:-1]
            print(len(_raw_data))
            double_quotes_position = Application.__find_all_substring_indices(_raw_data, '"')
            tab_position = Application.__find_all_substring_indices(_raw_data, '\t')
            line_break_position = Application.__find_all_substring_indices(_raw_data, '\n')

            valid_double_quotes_position = list()
            if double_quotes_position:
                _index, _end = 0, len(double_quotes_position)
                while _index < _end:
                    if _index == _end - 1:
                        valid_double_quotes_position.append(double_quotes_position[_index])
                        break
                    if double_quotes_position[_index] + 1 == double_quotes_position[_index + 1]:
                        _index += 2
                        continue
                    valid_double_quotes_position.append(double_quotes_position[_index])
                    _index += 1

            print("valid_double_quotes_position:")
            print(valid_double_quotes_position)

            if len(valid_double_quotes_position) % 2:
                raise Exception("error in valid_double_quotes_position")

            closed_areas = list()
            for index in range(len(valid_double_quotes_position) // 2):
                closed_areas.append((valid_double_quotes_position[index * 2:index * 2 + 2]))

            print("closed areas:")
            print(closed_areas)

            function_tab_counter = 0
            valid_separator_position = set()
            print("tab_position")
            print(tab_position)
            for position in tab_position:
                fail = False
                for area in closed_areas:
                    if area[0] <= position <= area[1]:
                        fail = True
                        break

                if not fail:
                    function_tab_counter += 1
                    valid_separator_position.add(position)

            print("function_tab_counter")
            print(function_tab_counter)

            for position in line_break_position:
                fail = False
                for area in closed_areas:
                    if area[0] <= position <= area[1]:
                        fail = True
                        break

                if not fail:
                    valid_separator_position.add(position)

            valid_separator_position.add(-1)
            valid_separator_position.add(len(_raw_data))
            valid_separator_position = sorted(list(valid_separator_position))
            print("valid_separator_position")
            print(valid_separator_position)

            _row = len(valid_separator_position) - 1 - function_tab_counter
            _col = (len(valid_separator_position) - 1) // _row
            print("[row, col]:")
            print(_row, _col)

            _counter = 0
            tmp_list = []
            for index in range(len(valid_separator_position) - 1):
                _counter += 1
                tmp_list.append(_raw_data[valid_separator_position[index] + 1:valid_separator_position[index + 1]])
                if _counter == _col:
                    ret_list.append(tmp_list)
                    _counter = 0
                    tmp_list = list()
            if tmp_list:
                ret_list.append(tmp_list)
        return ret_list

    def python2excel(self):
        input_content = self.output_area.get("1.0", END)
        tmp_ret = list()
        tmp = re.findall(r'\[(.*)]', input_content)
        if tmp:
            for _ in tmp:
                tmp_ret.append(list(eval(_)))
            print(tmp_ret)

            if self.data_direction.get() == 1:
                tmp_ret = list(zip(*tmp_ret))
            ret_str = ""
            for _ in tmp_ret:
                for __ in _[:-1]:
                    ret_str += str(__) + '\t'
                ret_str += str(_[-1]) + '\n'
            self.input_area.delete("1.0", END)
            self.input_area.insert("end", ret_str)
            self.input_area.insert(INSERT, '\n')
            if self.auto_copy_or_not.get() == 1:
                pyperclip.copy(self.input_area.get("1.0", END))
        pass

    def excel2python(self):
        with_header_flag = True if self.output_with_names_or_not.get() == 1 else False
        is_direction_vertical = True if self.data_direction.get() == 1 else False
        with_hierarchical_structure = True if self.has_hierarchical_structure_or_not.get() else False
        input_content = self.input_area.get("1.0", END)

        if input_content:
            self.output_area.delete("1.0", END)
            if not with_hierarchical_structure:
                if '\t' in input_content:
                    parsed_data = Application.__data_parser(input_content)
                    print(parsed_data)
                    if is_direction_vertical:
                        parsed_data = list(zip(*parsed_data))
                    for index, _ in enumerate(parsed_data):
                        self.output_area.insert("end", self.excel2python_print(_, with_header_flag))
                        self.output_area.insert(INSERT, '\n')
                else:
                    output_content = re.split(r'\n+', input_content)[:-1]
                    self.output_area.insert("end", self.excel2python_print(output_content, with_header_flag))
                    self.output_area.insert(INSERT, '\n')
            else:
                parsed_data = Application.__data_parser(input_content)
                hierarchical_data = list(zip(*parsed_data))

                aux_dict = dict()
                Application.__list_to_dict(hierarchical_data, aux_dict, len(hierarchical_data) - 2, len(hierarchical_data) - 2)
                self.output_area.insert("end", json.dumps(aux_dict, indent=4, ensure_ascii=False))
                self.output_area.insert(INSERT, '\n')
                pass

            if self.auto_copy_or_not.get() == 1:
                pyperclip.copy(self.output_area.get("1.0", END))

        pass

    @staticmethod
    def __list_to_dict(raw_data, record_dict, layer, last_layer):
        if layer == -1:
            return

        if last_layer == layer:
            cur_key = None
            for index, item in enumerate(raw_data[layer]):
                if item == "":
                    if cur_key:
                        record_dict[cur_key].append(raw_data[layer + 1][index])
                    else:
                        print("error1")
                        return
                else:
                    cur_key = item
                    record_dict[cur_key] = [raw_data[layer + 1][index], ]
        else:
            cur_key = None
            for index, item in enumerate(raw_data[layer]):
                if item != "":
                    cur_key = item
                    record_dict[cur_key]= dict()
                    try:
                        if raw_data[layer + 1][index]:
                            record_dict[cur_key][raw_data[layer + 1][index]] = record_dict[raw_data[layer + 1][index]]
                            del record_dict[raw_data[layer + 1][index]]
                        else:
                            record_dict[cur_key][raw_data[layer + 1][index]] = dict()
                    except:
                        pass
                elif raw_data[layer + 1][index]:
                    try:
                        record_dict[cur_key][raw_data[layer + 1][index]] = record_dict[raw_data[layer + 1][index]]
                        del record_dict[raw_data[layer + 1][index]]
                    except:
                        pass
                else:
                    pass

        print(record_dict)
        Application.__list_to_dict(raw_data, record_dict, layer - 1, last_layer)

    def excel2python_print(self, raw, need_var_name):
        ret = ''
        self.output_counter += 1
        if need_var_name:
            ret += self.output_default_prefix + str(self.output_counter) + ' = '
        ret += '['
        for _ in raw[:-1]:
            try:
                float(_)
            except:
                ret += '"' + str(_) + '",'
            else:
                if self.number_convert_to_string_or_not.get():
                    ret += '"' + str(_) + '",'
                else:
                    ret += str(_) + ', '

        try:
            float(raw[-1])
        except:
            ret += '"' + str(raw[-1]) + '"]'
        else:
            if self.number_convert_to_string_or_not.get():
                ret += '"' + str(raw[-1]) + '"]'
            else:
                ret += str(raw[-1]) + ']'
        return ret

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    A = Application()
    A.run()
    pass

