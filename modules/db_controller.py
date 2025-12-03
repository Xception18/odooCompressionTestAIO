import configparser
import logging

import psycopg2
import psycopg2.extras
import wx

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - [%(levelname)s] [%(threadName)s] (%(module)s:%(lineno)d) %(message)s",
    filename="aplikasiAlatUji.log",
)

import os
import sys

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

fileConfig = resource_path("config.cnf")
config = configparser.RawConfigParser()
config.read(fileConfig)
datab = config.get("data", "database")
hosted = config.get("data", "host")
login = config.get("data", "user")
passed = config.get("data", "password")


# Struktur tabel :
# idpengujian serial
# tgluji 	date 		0	self.txtTglPengujian
# idalat 	char(100) 	1	self.txtIdAlat
# kodebendaUjichar(200) 	2	self.txtKodeBendaUji
# nodocket 	char(100) 	3	self.txtNomorDocket
# nourutbenda 	char (2) 	4	self.txtNoUrut
# nilaikn 	numeric(10,2) 	5	self.txtHasilUji
# beratbenda 	numeric(10,2) 	6	self.txtBerat
# tiperetak 	char(1) 		7	self.chcTipeRetak
# sinkron 	char(1)
# idbendauji 	char(100) 	8	self.txtidBendaUji
# tglrencanauji date 		9	self.txtRencanaTglUji
# bujnama 	char(20)	10	self.txtJenisBUJ
# kuattekan 	numeric(10,2)	11	self.txtBeban
# bebanmpa 	numeric(10,2)	12	self.txtMpa
# umur 	integer		13	self.txtKg
# CONSTRAINT "idPengujian_PK" PRIMARY KEY ("idPengujian")


# Fungsi cekBendaUji(bendaUji) berfungsi untuk
def cekBendaUji(bendaUji):
    try:
        logging.info("Mulai mengeksekusi method cekBendaUji() pada file dbctrl.py")

        # ceknomerurut = []
        conn = f"dbname={datab} user={login} host={hosted} password={passed}"
        konekdb = psycopg2.connect(conn)
        konekdb.autocommit = True
        kursor = konekdb.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        data = (bendaUji,)
        SQL = """ SELECT nourutbenda FROM pengujian
			WHERE noDocket = %s ORDER BY nourutbenda DESC LIMIT 1 ; """
        kursor.execute(SQL, data)

        nomer = kursor.fetchall()
        # nomerBaru = int(nomer) + 1
        logging.debug("Data Benda uji hasil method cekBendaUji() : %s", str(nomer))
        print(" Nomer : ", nomer)
        # print " Nomer Baru : ", nomerBaru
        return nomer
            # return nomerBaru
            # print no
            # return no

    except Exception as e:
        logging.error(
            "Error pada saat menjalankan method cekBendaUji() pda file dbctrl.py",
            str(e),
        )
        str(e)
        print(e)
        # dlg = wx.MessageDialog(None, pesanError, "Error Koneksi Database", wx.OK | wx.ICON_INFORMATION)
        # dlg.ShowModal()
        # dlg.Destroy()


def simpan(bendaUji):
    try:
        print("bendaUji = ", bendaUji)
        logging.info("Mulai mengeksekusi method simpan() pada file dbctrl.py")
        conn = f"dbname={datab} user={login} host={hosted} password={passed}"
        konekdb = psycopg2.connect(conn)
        konekdb.autocommit = True
        kursor = konekdb.cursor()
        data = (
            bendaUji[0],
            bendaUji[1],
            bendaUji[2],
            bendaUji[3],
            bendaUji[4],
            bendaUji[5],
            bendaUji[6],
            bendaUji[7],
            "B",
            bendaUji[8],
            bendaUji[9],
            bendaUji[10],
            bendaUji[11],
            bendaUji[12],
            bendaUji[13],
            bendaUji[14],
        )

        print("data = ", data)
        SQL = """ INSERT INTO pengujian(tgluji, idalat, kodebendaUji, nodocket, nourutbenda, nilaikn, beratbenda, tiperetak, sinkron, idbendauji, tglrencanauji, bujnama, kuattekan, bebanmpa, umur, tglbendauji)
		VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s); """
        kursor.execute(SQL, data)
        logging.debug("Data berhasil disimpan")
        pesanError = "Data berhasil disimpan"
        dlg = wx.MessageDialog(
            None, pesanError, "Penyimpanan pada database", wx.OK | wx.ICON_INFORMATION
        )
        dlg.ShowModal()
        dlg.Destroy()
        kursor.close()
        konekdb.close()

            # sinkBendaUji(bendaUji)
    except Exception as e:
        logging.error(
            "Error pada saat menjalankan method simpan() pda file dbctrl.py : %s",
            str(e),
        )
        pesanError = str(e)
        # print e
        dlg = wx.MessageDialog(
            None, pesanError, "Error Koneksi Database", wx.OK | wx.ICON_INFORMATION
        )
        dlg.ShowModal()
        dlg.Destroy()


def queryBendaUji(bendaUji):
    try:
        logging.info("Mulai mengeksekusi method queryBendaUji() pada file dbctrl.py")
        conn = f"dbname={datab} user={login} host={hosted} password={passed}"
        konekdb = psycopg2.connect(conn)
        konekdb.autocommit = True
        kursor = konekdb.cursor()
        data = (
            bendaUji[0],
            bendaUji[1],
        )
        SQL = """ SELECT * FROM pengujian
		WHERE noDocket = %s AND noUrutBenda = %s) ORDER BY tgluji DESC; """
        kursor.execute(SQL, data)
        hasilSelect = kursor.fetchone()
        kursor.close()
        konekdb.close()
        logging.debug(
            "Data Benda uji hasil method queryBendaUji() : %s", str(hasilSelect)
        )
        return hasilSelect
    except Exception as e:
        logging.error(
            "Error pada saat menjalankan method queryBendaUji() pada file dbctrl.py",
            str(e),
        )
        pesanError = str(e)
        dlg = wx.MessageDialog(
            None, pesanError, "Error Koneksi Database", wx.OK | wx.ICON_INFORMATION
        )
        dlg.ShowModal()
        dlg.Destroy()


def queryGrid(parList):

    try:
        logging.info("Mulai mengeksekusi method queryGrid() pada file dbctrl.py")
        conn = f"dbname={datab} user={login} host={hosted} password={passed}"
        konekdb = psycopg2.connect(conn)
        konekdb.autocommit = True
        kursor = konekdb.cursor()
        dataList = (
            parList[0],
            parList[1],
            parList[2],
        )
        SQL = """ SELECT idpengujian, tgluji, nodocket, nourutbenda, bujnama, umur, nilaikn, bebanmpa, kuattekan, beratbenda, tiperetak, sinkron FROM pengujian
		WHERE  (tgluji BETWEEN %s AND %s) AND sinkron LIKE %s
		ORDER BY idpengujian ASC; """
        kursor.execute(SQL, dataList)
        hasilSelect = kursor.fetchall()
        logging.debug("Data Benda uji hasil method queryGrid() : %s", str(hasilSelect))
        kursor.close()
        konekdb.close()
        return hasilSelect
    except Exception as e:
        logging.error(
            "Error pada saat menjalankan method queryGrid() pda file dbctrl.py", str(e)
        )
        pesanError = str(e)
        dlg = wx.MessageDialog(
            None, pesanError, "Error Koneksi Database", wx.OK | wx.ICON_INFORMATION
        )
        dlg.ShowModal()
        dlg.Destroy()

        # SQL = """ SELECT idpengujian, tgluji, nodocket, nourutbenda, nilaikn, beratbenda tiperetak, sinkron FROM pengujian
        # WHERE tgluji >= %s AND tgluji <= %s
        # ORDER BY tgluji ASC; """
