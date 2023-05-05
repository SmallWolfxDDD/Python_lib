from itertools import permutations
import json
def allarr(string: str): return list(permutations(string)) 
class jsonx:
    def __init__(self, file_path):
        self.file_path = file_path
        
    def read(self, list_name=None):
        with open(self.file_path, 'r') as f:
            data = json.load(f)
            if list_name is not None:
                data = data[list_name]
        return data
    
    def write(self, data, list_name=None):
        with open(self.file_path, 'w') as f:
            json_data = json.load(f)
            if list_name is not None:
                json_data[list_name] = data
            else:
                json_data = data  
            json.dump(json_data, f)
