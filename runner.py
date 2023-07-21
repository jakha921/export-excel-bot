from export import open_files
from backet_parser import parse_backet


def run():
    parse_backet()
    open_files()
    print('ready/output.xlsx')
    return 'ready/output.xlsx'

if __name__ == '__main__':
    run()

