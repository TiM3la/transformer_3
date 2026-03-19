# classes.py

import math as m, csv, sqlite3 as q, os
import zipfile as z


# работа с базой данных
class DateBase:
    def __init__(self, exists=False ,db_name='db_1.db'):
        if exists is True:
            self.connection = q.connect(db_name)
            self.cursor = self.connection.cursor()
        else:
            self.connection = q.connect(f'data_bases/{db_name}')
            self.cursor = self.connection.cursor()
            self.tables = []

            print(f'Создана база данных "{db_name}"')

    def create_db_table(self,table_name, header_list, type_list=0):
        if type_list == 0:
            type_list = ['TEXT'] * len(header_list)
        query = f'CREATE TABLE IF NOT EXISTS {table_name} ('
        for i in range(len(header_list)):
            query += header_list[i].replace(' ', '_').replace('.', '_') + ' ' + type_list[i] + ', '
        query = query[:-2] + ')'

        print(f'Создана таблица {table_name}')

        self.cursor.execute(query)
        self.cursor.execute(f'CREATE INDEX IF NOT EXISTS idx_timestamp ON {table_name} (Timestamp_UTC)')

        self.connection.commit()

    def insert_db_table(self,table_name, header_list, datas_table):

        query = f'INSERT OR IGNORE INTO {table_name} ('
        for i in range(len(header_list)):
            query += header_list[i] + ', '
        query = query[:-2] + f') VALUES ({"?, " * len(header_list)}'
        query = query[:-2] + ')'
        self.cursor.executemany(query, comma_dot(datas_table, 'comma'))
        self.connection.commit()
        print(f'Загружены данные в таблицу {table_name}')

    def select_table(self, query):
        self.cursor.execute(query)
        return comma_dot(self.cursor.fetchall())

    def query(self, query):
        self.cursor.execute(query)
        self.connection.commit()
        print('Запрос выполнен')

# таблица, взятая из бд
class DbTable:
    def __init__(self, table):
        self.table = table
        self.converted = []
        self.modifided = []
        self.str_data = []
    # конвертируем в классы

    def convert_class(self, conversions):
        new_table = []
        for row in self.table:
            new_row = []
            for cls, indices in conversions:
                args = []
                for i in indices:
                    if isinstance(i, int):
                        args.append(row[i])
                    else:
                        args.append(i)
                new_row.append(cls(*args))
            new_table.append(new_row)
        self.converted = new_table

    # рассчитываем все величины
    def add_calculate_values(
            self, Snom, kt, Rt, Xt, Gt, Bt,
            kn_vn_a, kn_vn_w,
            koff,
            ku_hv,
            ki_hv,
            ki_lv
    ):
        new_table = []
        for row in self.converted:
            T = row[15]

            if koff:

                modul = row[1].modul / (1 + kn_vn_a)
                phase = row[1].phase_deg - kn_vn_w
                row[1] = ComplexValue(modul, phase) * ku_hv

                modul = row[2].modul / (1 + kn_vn_a)
                phase = row[2].phase_deg - kn_vn_w
                row[2] = ComplexValue(modul, phase) * ku_hv

                modul = row[3].modul / (1 + kn_vn_a)
                phase = row[3].phase_deg - kn_vn_w
                row[3] = ComplexValue(modul, phase) * ku_hv

            # создаем трансформатор
            A = Phase(row[4], row[1])
            B = Phase(row[5], row[2])
            C = Phase(row[6], row[3])

            vn = ThreePhaseSystem(A, B, C)

            a = Phase(row[11], row[8])
            b = Phase(row[12], row[9])
            c = Phase(row[13], row[10])

            nn = ThreePhaseSystem(a, b, c)

            trans_1 = Transformer(
                vn, nn, Snom, kt,
                kn_vn_a, kn_vn_w,
                row[15]
            )

            new_table.append([
                row[0], # 1
                row[7], # 1
                row[14], # 1
                row[1], # 2
                row[2], # 2
                row[3], # 2
                row[4], # 2
                row[5], # 2
                row[6], # 2
                row[8], # 2
                row[9], # 2
                row[10], # 2
                row[11], # 2
                row[12],
                row[13],
                trans_1.vn.u1,
                trans_1.vn.i1,
                trans_1.nn.u1,
                trans_1.nn.i1,
                trans_1.I_m,
                trans_1.U0,
                trans_1.Z_T_kat.real, # R
                trans_1.Z_T_kat.imag, # X
                trans_1.Ym_T.real * 1000000, # G мкСм
                -trans_1.Ym_T.imag * 1000000, # B мкСм

                trans_1.Z_1_G.real,  # R
                trans_1.Z_1_G.imag,  # X
                trans_1.Ym_G.real * 1000000,  # G мкСм
                -trans_1.Ym_G.imag * 1000000,  # B мкСм

                trans_1.vn.S.modul, # S vn
                trans_1.vn.S.real, # P vn
                trans_1.vn.S.imag, # Q vn
                trans_1.nn.S.modul,  # S nn
                trans_1.nn.S.real,  # P nn
                trans_1.nn.S.imag,  # Q nn
                trans_1.dS.modul, # dS
                trans_1.dP, # dP
                trans_1.dQ, # dQ
                trans_1.kz, # kz
                trans_1.kt, # kt
                trans_1.kpd, # kpd

                trans_1.Skz_T.modul, # кз в Т
                trans_1.Pkz_T,
                trans_1.Qkz_T,

                trans_1.Shh_T.modul,  # xx в Т
                trans_1.Phh_T,
                trans_1.Qhh_T,

                trans_1.Skz_G.modul,  # кз в Г
                trans_1.Pkz_G,
                trans_1.Qkz_G,

                trans_1.Shh_G.modul,  # xx в Г
                trans_1.Phh_G,
                trans_1.Qhh_G,

                (trans_1.Z_T_kat.real - Rt) / Rt * 100,  # погрешности в Т
                (trans_1.Z_T_kat.imag - Xt) / Xt * 100,
                (trans_1.Ym_T.real * 1000000 - Gt) / Gt * 100,
                (-trans_1.Ym_T.imag * 1000000 - Bt) / Bt * 100,

                (trans_1.Z_1_G.real - Rt) / Rt * 100,  # погрешности в Г
                (trans_1.Z_1_G.imag - Xt) / Xt * 100,
                (trans_1.Ym_G.real * 1000000 - Gt) / Gt * 100,
                (-trans_1.Ym_G.imag * 1000000 - Bt) / Bt * 100,

                row[15],
            ])
        self.modifided = new_table

    # конвертируем в text
    def convert_str(self):
        new_table = []
        for row in self.modifided:
            new_row = []
            for e in row:
                if type(e) is ComplexValue:
                    new_row.extend([e.modul, e.phase_deg])
                else:
                    new_row.append(e)
            new_table.append(list(map(lambda x: round(x, 4) if type(x) is float else x, new_row)))
        self.str_data = new_table

    def usred(self, dt):
        k = int(float(dt) / 0.02)
        print(k)
        data_transpose = list(zip(*self.table))
        new_table = [[], []]

        print(self.table)

        for i in range(len(self.table) // k):
            new_table[0].append(medium(data_transpose[0][i * k:i * k+k]))
            new_table[1].append(medium(data_transpose[1][i * k:i * k+k]))
        print(new_table)
        return new_table

# файл таблицы
class DataTable:
    def __init__(self, dir): # csv или zip-файл
        self.dir = dir
        self.file_name = dir.split('/')[-1].split('.')[0]
        self.header = []
        self.data = []

        if self.dir[-3:] == 'zip':
            with z.ZipFile(self.dir, mode='r') as file_zip:
                files = file_zip.namelist()
                for file in files:
                    print(file)
                    dir = file_zip.read(file).decode('utf-8').splitlines()
                    count = 0
                    reader_obj = csv.reader(dir)
                    for row in reader_obj:
                        if count == 0:
                            self.header = row
                        else:
                            self.data.append(row)
                        count += 1

# комплексное значение синусоидального сигнала (I, U)
class ComplexValue:
    def __init__(self, v1: float, v2: float, f: float = 50, ktr: float = 1, forma: str = 'pokaz'):
        v1 = float(v1)
        v2 = float(v2)
        f = float(f)

        if forma == 'alg':
            self.real = v1 * ktr
            self.imag = v2 * ktr
            self.modul = m.sqrt(self.real ** 2 + self.imag ** 2)
            self.phase_rad = m.atan2(self.imag, self.real)
            self.phase_deg = m.degrees(self.phase_rad)
        else:  # pokaz (показательная форма)
            self.modul = v1 * ktr
            self.phase_deg = v2
            self.phase_rad = m.radians(self.phase_deg)
            self.real = self.modul * m.cos(self.phase_rad)
            self.imag = self.modul * m.sin(self.phase_rad)

        self.comp = complex(self.real, self.imag)
        self.f = f
        self.w = 2 * m.pi * self.f

    def __add__(self, other):
        if isinstance(other, ComplexValue):
            new_real = self.real + other.real
            new_imag = self.imag + other.imag
            return ComplexValue(new_real, new_imag, f=self.f, forma='alg')
        elif isinstance(other, (int, float)):
            return ComplexValue(self.real + other, self.imag, f=self.f, forma='alg')
        raise TypeError(f"Unsupported type: {type(other)}")

    def __radd__(self, other):
        return self.__add__(other)

    def __sub__(self, other):
        if isinstance(other, ComplexValue):
            new_real = self.real - other.real
            new_imag = self.imag - other.imag
            return ComplexValue(new_real, new_imag, f=self.f, forma='alg')
        elif isinstance(other, (int, float)):
            return ComplexValue(self.real - other, self.imag, f=self.f, forma='alg')
        raise TypeError(f"Unsupported type: {type(other)}")

    def __rsub__(self, other):
        if isinstance(other, (int, float)):
            return ComplexValue(other - self.real, -self.imag, f=self.f, forma='alg')
        raise TypeError(f"Unsupported type: {type(other)}")

    def __mul__(self, other):
        if isinstance(other, ComplexValue):
            new_real = self.real * other.real - self.imag * other.imag
            new_imag = self.real * other.imag + self.imag * other.real
            return ComplexValue(new_real, new_imag, f=self.f, forma='alg')
        elif isinstance(other, (int, float)):
            return ComplexValue(self.real * other, self.imag * other, f=self.f, forma='alg')
        raise TypeError(f"Unsupported type: {type(other)}")

    def __rmul__(self, other):
        return self.__mul__(other)

    def __truediv__(self, other):
        if isinstance(other, ComplexValue):
            denom = other.real ** 2 + other.imag ** 2
            new_real = (self.real * other.real + self.imag * other.imag) / denom
            new_imag = (self.imag * other.real - self.real * other.imag) / denom
            return ComplexValue(new_real, new_imag, f=self.f, forma='alg')
        elif isinstance(other, (int, float)):
            return ComplexValue(self.real / other, self.imag / other, f=self.f, forma='alg')
        raise TypeError(f"Unsupported type: {type(other)}")

    def __rtruediv__(self, other):
        if isinstance(other, (int, float)):
            denom = self.real ** 2 + self.imag ** 2
            new_real = (other * self.real) / denom
            new_imag = (-other * self.imag) / denom
            return ComplexValue(new_real, new_imag, f=self.f, forma='alg')
        raise TypeError(f"Unsupported type: {type(other)}")

    def conjugate(self):
        return ComplexValue(self.real, -self.imag, f=self.f, forma='alg')

    def __str__(self):
        return f"{self.real:.4f} + j{self.imag:.4f} (|Z|={self.modul:.4f}, φ={self.phase_deg:.2f}°)"

# 1 фаза
class Phase:
    def __init__(self, I: ComplexValue, U:ComplexValue):
        self.I = I
        self.U = U

    # сдвиг фазы u относительно i в градусах
    def phase_shift(self):
        return self.U.phase_deg - self.I.phase_deg

    def __str__(self):
        return f'i = {self.I}, u = {self.U}'

# трехфазная система
class ThreePhaseSystem:
    def __init__(self, a: Phase, b: Phase, c: Phase):
        self.a = a
        self.b = b
        self.c = c

        self.i1 = 1 / 3 * (self.a.I + a_const * self.b.I + a_const * a_const * self.c.I)
        self.u1 = 1 / 3 * (self.a.U + a_const * self.b.U + a_const * a_const * self.c.U)

        u2 = 1 / 3 * (self.a.U + a_const * a_const * self.b.U + a_const * self.c.U)

        # коэффициент несимметрии напряжений по обратной последовательности
        self.k2 = u2.modul / self.u1.modul * 100
        self.S = self.u1 * self.i1.conjugate()

    def info(self):
        print(
            f'''
            a: {self.a}
            b: {self.b}
            c: {self.c}
            
            u1: {self.u1}
            i1: {self.i1}
            
            k2 = {self.k2} %
            S = {self.S.real, self.S.imag}
            S = {self.S}
            '''
        )

# трансформатор
class Transformer:
    def __init__(
            self, high_side: ThreePhaseSystem, low_side: ThreePhaseSystem, Snom, kt
    ):
        self.nn = low_side
        self.vn = high_side

        self.group = 1
        self.kt = self.vn.u1.modul / self.nn.u1.modul

        self.dS = self.vn.S - self.nn.S
        self.dP = self.vn.S.real - self.nn.S.real
        self.dQ = self.vn.S.imag - self.nn.S.imag
        self.kz = (self.nn.i1.modul * 100) / (((Snom * 1000000) * m.sqrt(3)) / (3 * 400))
        self.kpd = self.nn.S.real / self.vn.S.real * 100

        Uv = self.vn.u1 * self.group * ComplexValue(1, 30)
        Iv = self.vn.i1 * self.group * ComplexValue(1, 30)
        un = self.nn.u1 * kt
        i_n = self.nn.i1 / kt
        self.I_m = Iv - i_n

        # т-схема
        self.Z_1_T = (Uv - un) / (Iv + i_n)
        self.Ym_T = 1 / ((un * Iv + Uv * i_n) / (Iv * Iv - i_n * i_n))
        self.U0 = (un * Iv + Uv * i_n) / (Iv + i_n)

        self.Z_T_kat = self.Z_1_T * 2

        self.S1_T = self.vn.i1.modul * self.vn.i1.modul * (self.Z_1_T)
        self.P1_T = self.S1_T.real
        self.Q1_T = self.S1_T.imag

        self.S2_T = self.nn.i1.modul * self.nn.i1.modul * (self.Z_1_T / kt ** 2)
        self.P2_T = self.S2_T.real
        self.Q2_T = self.S2_T.imag

        self.Skz_T = self.S1_T + self.S2_T
        self.Pkz_T = self.Skz_T.real
        self.Qkz_T = self.Skz_T.imag

        self.Shh_T = self.dS - self.Skz_T
        self.Phh_T = self.Shh_T.real
        self.Qhh_T = self.Shh_T.imag

        # Г-схема
        self.Z_1_G = (Uv - un) / i_n
        self.Ym_G = 1 / (Uv / (Iv - i_n))

        self.Shh_G = self.I_m.modul * self.I_m.modul / self.Ym_G
        self.Phh_G = self.Shh_G.real
        self.Qhh_G = self.Shh_G.imag

        self.Skz_G = self.dS - self.Shh_G
        self.Pkz_G = self.Skz_G.real
        self.Qkz_G = self.Skz_G.imag

        # отклонения вычисленных параметров

    def __str__(self):
        return f'R = {self.Z_kat.real}, X = {self.Z_kat.imag}, G = {self.Ym.real}, B = {-self.Ym.imag}, k2_nn = {self.nn.k2}, k2_vn = {self.vn.k2}, dP = {self.dP}, dQ = {self.dQ}, kz = {self.kz}, U0 = {self.U0}'

class KatalogData:
    def __init__(self, S, Uvn, Unn, Uk, dPk, Phh, Ih):
        self.dQh = Ih * S * 10 # квар
        self.Rt = dPk * Uvn * Uvn / 1000 / S / S
        self.Xt = Uk * Uvn * Uvn / 100 / S
        self.Gt = Phh / 1000 / Uvn / Uvn * 1000000
        self.Bt = self.dQh / 1000 / Uvn / Uvn * 1000000
        self.Kt = Uvn / Unn

# в текущем директории находит каждую таблицу csv.zip
def find_files(directory, extension='zip'):
    extension = extension.lstrip('.').lower()
    matched_files = []
    abs_directory = os.path.abspath(directory)

    for root, _, files in os.walk(abs_directory):
        for filename in files:
            file_ext = os.path.splitext(filename)[1].lower()
            if file_ext == f".{extension}":
                full_path = os.path.join(root, filename)
                matched_files.append(full_path)

    return matched_files

# берет нужные индексы и добавляет значения
def format_data(table, indexes, *args):
    new_list = []
    for i in table:
        list_1 = []
        for j in indexes:
            list_1.append(i[j])
        list_1.extend(args)
        new_list.append(list_1)
    return new_list

def comma_dot(data_table, value='dot'):
    for j in range(len(data_table)):
        data_table[j] = list(data_table[j])
        for i in range(len(data_table[j])):
            if '/' not in str(data_table[j][i]):
                if value == 'comma':
                    data_table[j][i] = str(data_table[j][i]).replace('.', ',')
                elif type(data_table[j][i]) is str:
                    data_table[j][i] = str(data_table[j][i]).replace(',', '.')
    return data_table
# комплексная константа
a_const = ComplexValue(1, 120)

# красивое время
def nice_t(t): # секунд
    if t / 60 < 1:
        return f'{round(t, 2)} c'
    else:
        return f'{round(t / 60, 2)} мин'

# усредняет или вычленяет:
def medium(list_1):
    if ' ' in list_1[0]: # если время
        print(len(list_1), list_1)
        return list_1[int(len(list_1) - 1 / 2)]
    else:
        print(len(list_1), list_1)
        return  sum(list(map(float, list_1))) / len(list_1)