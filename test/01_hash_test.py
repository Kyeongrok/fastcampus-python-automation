class HashTable:
    size = 10
    table = [[] for _ in range(size)]

    def name_hash(self, s):
        ascii_sum = 0
        for c in s:
            ascii_sum += ord(c)
        return ascii_sum % self.size

    def insert(self, key, value):
        idx = self.name_hash(key)
        self.table[idx].append({'key':key, 'value':value})

    def find(self, key):
        idx = self.name_hash(key)
        if len(self.table[idx]) == 1:
            return self.table[idx][0]['value']
        else:
            for item in self.table[idx]:
                if item['key'] == key:
                    return item['value']

    def print_table(self):
        print(self.table)



if __name__ == '__main__':
    ht = HashTable()
    ht.insert('jisoo', 11)
    ht.insert('sooji', 12)
    ht.print_table()
    print(ht.find('sooji'))
