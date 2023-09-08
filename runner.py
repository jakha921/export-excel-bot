from export import open_files
from backet_parser import parse_backet


def run():
    output_file_path, errors = parse_backet()
    if errors:
        return output_file_path, errors
    path, errors = open_files()
    print('path', path)
    return path, errors


if __name__ == '__main__':
    run()
