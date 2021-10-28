

class OnceError(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return "Variable {0} has been be assigned".format(repr(self.value))
