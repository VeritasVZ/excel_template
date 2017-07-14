import json

class Serializer:
    def __init__(self,filename = 'utils\lang_attribures.json'):
        self.filename = filename

    def load(self):
        try:
            with open(self.filename, 'rt', encoding='ascii') as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def save(self, data):
        with open(self.filename, 'wt', encoding='ascii') as f:
            json.dump(data,f)