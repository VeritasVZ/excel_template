import json
from BMS_template.json_serializer import Serializer

class LangTypeAttributes:
    def __init__(self, serializer):
        self.serializer = serializer
        self.attributes = self.serializer.load().copy()

    def create_attribute(self, attribute_set): #attributes takes tuple of 3-langs types (UA, RU, EN)
        if attribute_set[0] not in self.attributes and attribute_set[1] not in self.attributes and attribute_set[2] not in self.attributes:
            self.attributes[attribute_set[2]] = attribute_set
            self.serializer.save(self.attributes)
        else:
            print('Entry exists')
            pass

    def read_attribute(self, key):
        try:
            return self.attributes[key]
        except KeyError:
            raise ValueError('Entry exists')

    def read_all(self):
        try:
            return self.attributes.values()
        except KeyError:
            raise ValueError('Something wrong')

class AttributesToTemplate:
    def add_attributes(data): # data >>> dictionary with tuples of 3-langs types (UA, RU, EN)
        serializer = Serializer()
        attributes = LangTypeAttributes(serializer)
        items = list(data.keys())
        for item_no in range(len(items)):
            attribute_set = data.get(items[item_no])
            attributes.create_attribute(attribute_set=attribute_set)
            print(attributes.read_attribute(attribute_set[2]))

    def get_attributes(template_lang, keys_list): # takes integer arg 0,1 or 2 that refers to land UA, RU or EU(en)
        serializer = Serializer()
        attributes = LangTypeAttributes(serializer)
        json_data = attributes.read_all()
        #print('count of entries in doc: '+str(len(json_data)))
        template_attributes = []
        for item in json_data:
            if item[2] in keys_list:
                #print(item[template_lang])
                template_attributes.append(item[template_lang])
        return template_attributes