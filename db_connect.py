import sqlite3

class REPORT_DB():
    def __init__(self, date_base):
        """Подключаемся к Базе Данных и сохраняем курсор соединения"""
        self.connection = sqlite3.connect(date_base)
        self.cursor = self.connection.cursor()

    def add_progect_name(self, prog_name:str):
        """Добавляем проект в таблицу"""
        return self.cursor.execute("""INSERT INTO progect_main (progect_name) VALUES (?)""",
                                   (prog_name,)) and self.connection.commit()
    def delit_progect(self, progect_name=str):
        """Удаляем информацию"""
        result = self.cursor.execute("""DELETE FROM progect_main WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def get_id_progect(self, progect_name=str):
        """ПОлучаем id проекта"""
        self.cursor.execute("""SELECT id FROM progect_main WHERE progect_name = ? """, (progect_name,))
        return self.cursor.fetchall()

    def get_all_info_progect(self):
        """Получаем всю информацию из таблицы проектов"""
        self.cursor.execute("""SELECT * FROM progect_main""")
        return self.cursor.fetchall()

    def get_info_about_this_progect(self, progect=str):
        """Получаем всю информацию из таблицы проектов"""
        self.cursor.execute("""SELECT * FROM progect_main WHERE progect_name = ?""", (progect,))
        return self.cursor.fetchall()

    def get_info_progect(self):
        """Получаем информацию о проектах"""
        self.cursor.execute("""SELECT * FROM progect_main""")
        return self.cursor.fetchall()

    def get_info_progect_kmd(self):
        """Получаем информацию о проектах"""
        self.cursor.execute("""SELECT progect_name FROM progect_main WHERE kmd_active = TRUE""")
        return self.cursor.fetchall()

    def get_info_progect_tent(self):
        """Получаем информацию о проектах"""
        self.cursor.execute("""SELECT progect_name FROM progect_main WHERE tent_common_active = TRUE""")
        return self.cursor.fetchall()

    def get_info_progect_spu(self):
        """Получаем информацию о проектах"""
        self.cursor.execute("""SELECT progect_name FROM progect_main WHERE spu_active = TRUE""")
        return self.cursor.fetchall()

    def get_info_progect_pvh(self):
        """Получаем информацию о проектах"""
        self.cursor.execute("""SELECT progect_name FROM progect_main WHERE pvh_active = TRUE""")
        return self.cursor.fetchall()

    def get_info_progect_montage(self):
        """Получаем информацию о проектах"""
        self.cursor.execute("""SELECT progect_name FROM progect_main WHERE montage_active = TRUE""")
        return self.cursor.fetchall()

    def update_active_kmd(self, progect_name:str):
        """Обновляем Актуальность проекта КМД"""
        result = self.cursor.execute("""UPDATE progect_main SET kmd_active = TRUE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_disactive_kmd(self, progect_name:str):
        """Добавляем в Архив проект КМД"""
        result = self.cursor.execute("""UPDATE progect_main SET kmd_active = FALSE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_active_tent(self, progect_name:str):
        """Проект имеет и СПУ и ПВХ"""
        result = self.cursor.execute("""UPDATE progect_main SET tent_common_active = TRUE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_disactive_tent(self, progect_name:str):
        """Проект не имеет либо СПУ и ПВХ"""
        result = self.cursor.execute("""UPDATE progect_main SET tent_common_active = FALSE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_active_spu(self, progect_name:str):
        """Обновляем Актуальность проекта СПУ"""
        result = self.cursor.execute("""UPDATE progect_main SET spu_active = TRUE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_disactive_spu(self, progect_name:str):
        """Добавляем в архив проект СПУ"""
        result = self.cursor.execute("""UPDATE progect_main SET spu_active = FALSE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_active_pvh(self, progect_name: str):
        """Обновляем Актуальность проекта ПВХ"""
        result = self.cursor.execute("""UPDATE progect_main SET pvh_active = TRUE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_disactive_pvh(self, progect_name: str):
        """Добавляем в архив проект ПВХ"""
        result = self.cursor.execute("""UPDATE progect_main SET pvh_active = FALSE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_progect_name(self, progect_name:str, id:int):
        """Обновляем данные в таблице"""
        result = self.cursor.execute("""UPDATE progect_main SET progect_name = ? WHERE id = ?""",
                                     (progect_name, id,)) and self.connection.commit()
        return result

    def get_massa_progect(self):
        """ПОлучаем Массу КМД проекта"""
        self.cursor.execute("""SELECT progect_name, key_progect FROM progect_main WHERE kmd_active = TRUE """)
        return self.cursor.fetchall()

    def get_squer_spu_progect(self):
        """ПОлучаем Площадь СПУ проекта"""
        self.cursor.execute("""SELECT progect_name, key_progect FROM progect_main WHERE spu_active = TRUE """)
        return self.cursor.fetchall()

    def get_squer_pvh_progect(self):
        """ПОлучаем Площадь ПВХ проекта"""
        self.cursor.execute("""SELECT progect_name, key_progect FROM progect_main WHERE pvh_active = TRUE """)
        return self.cursor.fetchall()

    def delite_prod(self, progect: str):
        """Удаляем строку в таблице """
        result = self.cursor.execute("""DELETE FROM progect_main WHERE progect_name = ?""", (progect,)) and self.connection.commit()
        return result

    def add_reporting_data(self, prog_name=str, date='dd MMM yy', blank_per=str, weld_per=str, print_per=str,
                           common_report_per=str, ready_massa=str, kmd_prog_massa=str, key=str, real_date=str):
        """Добавляем данные о изготовлении ангара в таблицу КМД"""
        return self.cursor.execute("""INSERT INTO report_for_progect_2 (progect_name, data, blank_perc,
         weld_perc, print_perc, report_ready_perc, kmd_prog_massa, ready_massa, key_progect, real_date)
         VALUES (?,?,?,?,?,?, ?,?,?,?)""", (prog_name, date, blank_per, weld_per, print_per,
                                          common_report_per, kmd_prog_massa, ready_massa, key, real_date,)) and self.connection.commit()

    def get_info_ab_day_report(self, prog_name):
        """Достаем всю информацию из таблицы готовности КМД"""
        self.cursor.execute("""SELECT * FROM report_for_progect_2 WHERE progect_name = ? ORDER BY real_date ASC""", (prog_name,))
        return self.cursor.fetchall()

    def get_info_about_today_kmd_report(self, key, date = 'dd MMM yy'):
        """Достаем всю информацию из таблицы готовности КМД"""
        self.cursor.execute("""SELECT progect_name, kmd_prog_massa, ready_massa, report_ready_perc
         FROM report_for_progect_2 WHERE key_progect = ? and data = ? ORDER BY real_date ASC""", (key, date,))
        return self.cursor.fetchall()

    def delit_data_report(self, progect_name=str, data=str):
        """Удаляем информацию о КМД"""
        result = self.cursor.execute("""DELETE FROM report_for_progect_2 WHERE progect_name = ? and data = ?""", (progect_name, data,)) and self.connection.commit()
        return result

    def add_report_spu(self, prog_name=str, date='dd-mm-yyyy', raskroy_polypr=str, stitching=str, gloe_polypr=str,
                           punch_luverc=str, ready_prod = str, common_proc = str, spu_prog_sq=str, ready_squer_spu=str,
                       key=str, real_date=str):
        """Добавляем данные о изготовлении ангара в таблицу СПУ"""
        return self.cursor.execute("""INSERT INTO report_for_spu (progect_name, data_spu, raskroy_polypr,
         stitching, gloe_polypr, punch_luverc, ready_prod, common_proc, spu_prog_sq, ready_squer_spu, key_progect, real_date)
         VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""", (prog_name, date, raskroy_polypr, stitching,
                                             gloe_polypr, punch_luverc, ready_prod, common_proc,spu_prog_sq,
                                             ready_squer_spu, key, real_date)) and self.connection.commit()

    def get_info_ab_day_report_spu(self, prog_name):
        """Достаем всю информацию из таблицы готовности СПУ"""
        self.cursor.execute("""SELECT * FROM report_for_spu WHERE progect_name = ? ORDER BY real_date ASC""", (prog_name,))
        return self.cursor.fetchall()

    def get_tent_day_report_spu(self, prog_name, date):
        """Достаем всю информацию из таблицы готовности  по дате"""
        self.cursor.execute("""SELECT * FROM report_for_spu WHERE progect_name = ? AND data_spu = ? ORDER BY real_date ASC""", (prog_name, date,))
        return self.cursor.fetchall()

    def get_info_about_today_spu_report(self, key, date = 'dd MMM yy'):
        """Достаем всю информацию из таблицы готовности КМД"""
        self.cursor.execute("""SELECT progect_name, spu_prog_sq, ready_squer_spu, common_proc, ready_prod
         FROM report_for_spu WHERE key_progect = ? and data_spu = ?""", (key, date,))
        return self.cursor.fetchall()


    def delit_report_spu_data(self, progect_name=str, data=str):
        """Удаляем информацию о СПУ"""
        result = self.cursor.execute("""DELETE FROM report_for_spu WHERE progect_name = ? and data_spu = ?""", (progect_name, data,)) and self.connection.commit()
        return result

    def add_report_pvh(self, prog_name=str, date='dd-mm-yyyy', raskroy_polotna=str, raskroy_pockets=str, nashelnik=str, weld_pockets=str, weld_in_pockets=str, stitching=str,
                           weld_nashelnik=str, ready_prod=str, common_proc=str, pvh_prog_sq=str, ready_squer_pvh=str, key=str, real_date=str):
        """Добавляем данные о изготовлении ангара в таблицу ПВХ"""
        return self.cursor.execute("""INSERT INTO report_for_pvh (progect_name, data, raskroy_polotna, raskroy_pockets, nashelnik,
         weld_pockets, weld_in_pockets, stitching, weld_nashelnik, ready_prod, common_proc, pvh_prog_sq, ready_squer_pvh, key_progect, real_date)
         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (prog_name, date, raskroy_polotna, raskroy_pockets,
                                                   nashelnik, weld_pockets, weld_in_pockets, stitching,
                                                   weld_nashelnik, ready_prod, common_proc, pvh_prog_sq,
                                                   ready_squer_pvh, key, real_date,)) and self.connection.commit()

    def get_info_ab_day_report_pvh(self, prog_name):
        """Достаем всю информацию из таблицы готовности ПВХ"""
        self.cursor.execute("""SELECT * FROM report_for_pvh WHERE progect_name = ? ORDER BY real_date ASC""", (prog_name,))
        return self.cursor.fetchall()

    def get_info_about_today_pvh_report(self, key, date = 'dd MMM yy'):
        """Достаем всю информацию из таблицы готовности КМД"""
        self.cursor.execute("""SELECT progect_name, pvh_prog_sq, ready_squer_pvh, common_proc, ready_prod FROM report_for_pvh WHERE key_progect = ? and data = ?""", (key, date,))
        return self.cursor.fetchall()

    def delit_report_pvh_data(self, progect_name=str, data=str):
        """Удаляем информацию о ПВХ"""
        result = self.cursor.execute("""DELETE FROM report_for_pvh WHERE progect_name = ? and data = ?""", (progect_name, data,)) and self.connection.commit()
        return result

    def add_plan_for_mounth(self, date, plan_for_kmd, plan_for_tent, plan_spu, data_real):
        """Добавляем план на месяц в таблицу"""
        return self.cursor.execute("""INSERT INTO mounth_plan (mounth_year, plan_kmd, plan_tent, plan_spu, data_real)
         VALUES (?, ?, ?, ?, ?)""",
                                   (date, plan_for_kmd, plan_for_tent, plan_spu, data_real,)) and self.connection.commit()
    def delite_plan(self, date=str):
        """Удаляем информацию о плане"""
        result = self.cursor.execute("""DELETE FROM mounth_plan WHERE mounth_year = ?""", (date,)) and self.connection.commit()
        return result

    def update_plan_for_mounth(self,plan_kmd, plan_tent, plan_spu, date):
        """Обновляем Актуальность проекта ПВХ"""
        result = self.cursor.execute("""UPDATE mounth_plan SET plan_kmd = ? AND plan_tent = ? AND plan_spu = ?WHERE data_real = ?""",
                                     (plan_kmd, plan_tent, plan_spu, date,)) and self.connection.commit()
        return result

    def get_real_data_for_plan(self, date):
        """ПОлучаем значение реальной даты"""
        self.cursor.execute("""SELECT * FROM mounth_plan WHERE mounth_year = ? """, (date,))
        return self.cursor.fetchall()

    def get_all_plan_for_every_mounth(self):
        """ПОлучаем все из таблица плановой"""
        self.cursor.execute("""SELECT * FROM mounth_plan """)
        return self.cursor.fetchall()

    def get_kmd_plan_for_mounth(self, date):
        """ПОлучаем значение плана для кмд"""
        self.cursor.execute("""SELECT plan_kmd FROM mounth_plan WHERE mounth_year = ? """, (date,))
        return self.cursor.fetchone()

    def get_tent_plan_for_mounth(self, date):
        """ПОлучаем значение плана для тентового цеха"""
        self.cursor.execute("""SELECT plan_tent FROM mounth_plan WHERE mounth_year = ? """, (date,))
        return self.cursor.fetchone()

    def get_spu_plan_for_mounth(self, date):
        """ПОлучаем значение плана для тентового цеха"""
        self.cursor.execute("""SELECT plan_spu FROM mounth_plan WHERE mounth_year = ? """, (date,))
        return self.cursor.fetchone()

    def update_active_montage(self, progect_name: str):
        """Обновляем Актуальность монтажа"""
        result = self.cursor.execute("""UPDATE progect_main SET montage_active = TRUE WHERE progect_name = ?""",
                                     (progect_name,)) and self.connection.commit()
        return result

    def update_disactive_montage(self, progect_name: str):
        """Добавляем в архив монтаж"""
        result = self.cursor.execute("""UPDATE progect_main SET montage_active = FALSE WHERE progect_name = ?""",
               (progect_name,)) and self.connection.commit()
        return result

    def get_all_active_montage_progect(self):
        """ПОлучаем все активные проекты где ведется мотаж"""
        self.cursor.execute("""SELECT progect_name, key_progect FROM progect_main WHERE montage_active = TRUE """)
        return self.cursor.fetchall()

    def add_montag_everyday_report(self, progect_name, date, org_work, instl_metal_frame,instl_fenc_constr, instl_engin_syst,
                                   finish_work, com_perc_of_work, probl, solve_probl, real_date, key):
        """Добавляем данные о изготовлении ангара в таблицу Монтажа"""
        return self.cursor.execute("""INSERT INTO report_for_montage (progect_name, date, org_work, instal_of_metal_frame,
         instal_of_fencing_constr, instal_of_engineer_syst, finish_work, common_pecr_of_work, problems,
          way_to_solve_problems, real_date, key_progect)
         VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""", (progect_name, date, org_work, instl_metal_frame,instl_fenc_constr,
                                               instl_engin_syst, finish_work, com_perc_of_work,
                                               probl, solve_probl, real_date, key, )) and self.connection.commit()

    def delite_data_montag_report(self, progect, date):
        """Удаляем информацию о монтаже"""
        result = self.cursor.execute("""DELETE FROM report_for_montage WHERE key_progect = ? and real_date = ?""",
                                     (progect, date,)) and self.connection.commit()
        return result

    def delite_data_montag_report_by_date(self, progect, date):
        """Удаляем информацию о монтаже"""
        result = self.cursor.execute("""DELETE FROM report_for_montage WHERE key_progect = ? and date = ?""",
                                     (progect, date,)) and self.connection.commit()
        return result

    def get_info_evr_report_montage(self, prog_name):
        """Достаем всю информацию из таблицы готовности ПВХ"""
        self.cursor.execute("""SELECT * FROM report_for_montage WHERE progect_name = ? ORDER BY real_date ASC""", (prog_name,))
        return self.cursor.fetchall()
    
    def get_info_report_montage(self, prog_name, real_date):
        """Достаем всю информацию из таблицы готовности ПВХ"""
        self.cursor.execute("""SELECT * FROM report_for_montage WHERE progect_name = ? and real_date = ? ORDER BY real_date ASC""", (prog_name, real_date,))
        return self.cursor.fetchall()

    def get_info_about_today_montage_report(self, key, date = 'dd MMM yy'):
        """Достаем всю информацию из таблицы готовности КМД"""
        self.cursor.execute("""SELECT progect_name, common_pecr_of_work FROM report_for_montage WHERE key_progect = ? and date = ?""", (key, date,))
        return self.cursor.fetchall()

    def change_type_of_menu_produce(self):
        """Меняем меню на производственное"""
        result = self.cursor.execute("""UPDATE options_table SET menu_choose = 2 WHERE id = 1""") and self.connection.commit()
        return result

    def change_type_of_menu_montage(self):
        """Меняем меню на монтажное подразделение"""
        result = self.cursor.execute("""UPDATE options_table SET menu_choose = 3 WHERE id = 1""") and self.connection.commit()
        return result

    def change_type_of_menu_common(self):
        """Общее меню"""
        result = self.cursor.execute("""UPDATE options_table SET menu_choose = 1 WHERE id = 1""") and self.connection.commit()
        return result

    def get_menu_main(self):
        """Достаем вид меню"""
        self.cursor.execute("""SELECT menu_choose FROM options_table WHERE id = 1""")
        return self.cursor.fetchall()