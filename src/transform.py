# coding=utf-8
import csv
from glob import glob
from docx import Document


def extract_text(filename):
    doc = Document(filename)
    text = [p.text for p in doc.paragraphs]
    return [r for r in text if r]


def parse_line(line):
    line = line.encode('utf-8').replace('—', '-')
    key, val = line.split(':', 1)
    if '-' in key:
        start, end = [s.lower() for s in key.split('-')]
    else:
        start = end = key.lower()
    return [s.strip() for s in [start, end, val]]


def generate_csv():
    data_files = glob('../data/*.docx')
    with open('../data/combined.csv', 'wb') as fp:
        csv_writer = csv.writer(fp, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
        csv_writer.writerow(['id', 'start', 'end', 'bio'])
        for filename in data_files:
            text = extract_text(filename)
            uid = filename.split('_')[-1].split('.')[0]
            for line in text:
                csv_writer.writerow([uid] + parse_line(line))


if __name__ == '__main__':
    generate_csv()
