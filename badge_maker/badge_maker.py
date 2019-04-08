from __future__ import print_function
import re
import argparse
import pandas as pandas
import os
from tqdm import tqdm
from functools import partial
import numpy as np

template_keys = ["#VORNAME", "#NACHNAME", "#AUFGABE"]
regexes = {key: "({})+".format(key) for key in template_keys}


def main(excel_file, template_svg):
    def conv_func(input):
        output = str(input)
        return output
    df = pandas.read_excel(excel_file, converters={key : conv_func for key in template_keys}, na_filter=False)
    value_dict = { key: [row[key] for _, row in df.iterrows()] for key in template_keys }
    pointer_dict = { key: 0 for key, _ in value_dict.items() }

    #prepare current outile
    i_outfile = 0
    done = False

    while not done:
        i_outfile += 1
        with open(next_out_file(i_outfile, template_svg), 'w') as output_file:
            with open(template_svg, 'r') as template_file:
                for line in template_file:
                    new_line = line

                    # search line and replace
                    for key_to_replace, regex in regexes.items():
                        cur_pointer = pointer_dict[key_to_replace]
                        value_list = value_dict[key_to_replace]
                        if cur_pointer < len(value_list):
                            cur_value = str(value_list[cur_pointer])
                            if re.search(regex, new_line):
                                if len(cur_value) == 0:
                                    cur_value = str("")

                                print(str("{} {} {}").format(key_to_replace, cur_pointer, cur_value))
                                new_line = re.sub(regex, cur_value, new_line)
                                pointer_dict[key_to_replace] += 1
                        else:
                            pass
                    # print new line
                    print(new_line.encode("utf-8"), file=output_file)
                    if np.all([len(lst) == pointer_dict[key] for key, lst in value_dict.items()]):
                        done = True
                    

        if cur_pointer == 0:
            print("no matches were found. please check your options again")
            done = True
    
    summary(i_outfile, pointer_dict)


def next_out_file(i_outfile, template_svg):
    print('creating new output file number {}'.format(i_outfile))
    path, ext = os.path.splitext(template_svg)
    path = "{}_output_{}".format(path, i_outfile)
    output_svg = ''.join([path, ext])
    return output_svg


def summary(i_outfile, pointer_dict):
    print("---Summary---")
    print("number of output files created : {}".format(i_outfile))
    template = "number of {} inserted : {}"
    for key, count_value in pointer_dict.items():
        print(template.format(key, count_value))


if __name__ == '__main__':
    import argparse
    epilog = """
    SVG generator. Finds template values within the SVG and replaces them with the fields from the excel file, one at a time.
    Template keys:
        template_keys = ["#VORNAME", "#NACHNAME", "#AUFGABE"]

    If there are more fields in the excel table than the template has fields, multiple output files will be generated

    example usage:
        python badge_maker.py --excel-file example_table.xlsx --template-svg template_grid.svg


    Note: excel file values must not contain & characters or other characters that need to be escaped.
    """
    parser = argparse.ArgumentParser(epilog=epilog)
    parser.add_argument('--excel-file', default='example.xlsx', help="path to .xlsx-file")
    parser.add_argument('--template-svg', default='example_template.svg', help="template.svg")
    args = parser.parse_args()
    main(args.excel_file, args.template_svg)