import configparser


def load_config(file):
    config = configparser.ConfigParser()
    config.read(file)
    return {
        'drawing': {
            'number_digits': config.getint('drawing', 'number_digits', fallback=3)
        }
    }


if __name__ == '__main__':
    pass
