import csv
import codecs


def writer(filename, data, header):
    with open(filename, 'w', encoding='utf-8', newline='') as my_file:
        csv_writer = csv.writer(my_file)
        csv_writer.writerow(header)
        csv_writer.writerows(data)
    my_file.close()


def append(filename, data, header):
    with open(filename, 'a', encoding='utf-8', newline='') as my_file:
        csv_writer = csv.writer(my_file)
        csv_writer.writerow(header)
        csv_writer.writerows(data)
    my_file.close()


def read(myfile):
    my_list = []
    with open(myfile, 'r') as my_csv:
        csv_reader = csv.reader(my_csv)
        #next(csv_reader)
        for item in csv_reader:
            my_list.append(item)
        my_csv.close()
    return my_list
