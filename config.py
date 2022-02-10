import configparser


def load_config(file=None):
    config = configparser.ConfigParser()
    if file is not None:
        config.read(file)
    return {
        'drawing': {
            'number_digits': config.getint(section='drawing', option='number_digits', fallback=4),
            'unit_digits': config.getint(section='drawing', option='unit_digits', fallback=2),
            'number_prefix': config.get(section='drawing', option='number_prefix', fallback=''),
            'start_unit': config.getint(section='drawing', option='start_unit', fallback=1),
        },
        'connector': {
            'number_digits': config.getint(section='connector', option='number_digits', fallback=6),
        }
    }


if __name__ == '__main__':
    print(load_config())
