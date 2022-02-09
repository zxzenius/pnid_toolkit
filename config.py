import configparser


class Project:
    __conf = None

    @staticmethod
    def config():
        if Project.__conf is None:
            Project.__conf = configparser.ConfigParser()
            Project.__conf.read('config.ini')
        return Project.__conf


if __name__ == '__main__':
    print(Project.config()['connector']['prefix_digits'])
