from export import open_files
from backet_parser import parse_backet


def run():
    parse_backet()
    path = open_files()
    print('path', path)
    return path


if __name__ == '__main__':
    run()
