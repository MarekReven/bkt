def create_db():
    conn = sqlite3.connect("database.db")
    curs = conn.cursor()
    try:
        curs.execute("CREATE TABLE IF NOT EXISTS voltage_tbl (voltage_code TEXT, voltage_desc TEXT);")
    except Exception as e:
        raise e
    return conn, curs

def insert_data_to_sqlite(conn, curs):

    reader = csv.reader(open('data_table.csv', 'r'), delimiter=',')
    curs.execute("DELETE FROM voltage_tbl;")
    try:
        for row in reader:
            to_db = [row[0], row[1]]
            curs.execute("INSERT INTO voltage_tbl (voltage_code, voltage_desc) VALUES (?, ?);", to_db)
        conn.commit()
    except Exception as e:
        raise e
    return conn

def make_voltage_desc_list(conn):
    c = conn.cursor()
    kody = []
    for row in c.execute('SELECT voltage_code FROM voltage_tbl'):
        kody.append(str(row).strip("(),'"))
    return kody

def find_voltage_code(code):
    conn = sqlite3.connect("database.db")
    curs = conn.cursor()
    voltage_desc = curs.execute('SELECT voltage_code FROM voltage_tbl WHERE voltage_code =' + str(code))
    print(voltage_desc)

# n = [1,2,3]
# l = ['a', 'b', 'c']
#
# for x,y in zip(n,l):
#     print(x, y)

l = [1,2,3]
print(len(l))

