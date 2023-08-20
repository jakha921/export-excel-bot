from export import open_files
from backet_parser import parse_backet


def run():
    parse_backet()
    path, errors = open_files()
    print('path', path)
    return path, errors


if __name__ == '__main__':
    run()
