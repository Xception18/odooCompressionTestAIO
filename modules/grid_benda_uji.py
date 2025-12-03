import datetime
import logging
import sys
import os
import time

# Add project root to sys.path so we can import from modules
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import wx
import wx.adv

from modules import db_controller as dbctrl

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - [%(levelname)s] [%(threadName)s] (%(module)s:%(lineno)d) %(message)s",
    filename="grid_benda_uji.log",
)


class GridBendaUji(wx.Panel):
    def __init__(self, parent):
        logging.info("Inisialisasi Aplikasi Pengendali Alat Penguji Tekanan")
        wx.Panel.__init__(
            self,
            parent,
            id=wx.ID_ANY,
            pos=wx.DefaultPosition,
            size=wx.Size(810, 423),
            style=wx.TAB_TRAVERSAL,
        )

        sizerBoxUtama = wx.BoxSizer(wx.VERTICAL)

        sizerFlexSort = wx.FlexGridSizer(0, 7, 0, 0)
        sizerFlexSort.SetFlexibleDirection(wx.BOTH)
        sizerFlexSort.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.tglAwal = wx.adv.DatePickerCtrl(
            self,
            wx.ID_ANY,
            wx.DefaultDateTime,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.adv.DP_DEFAULT | wx.adv.DP_SHOWCENTURY,
        )
        sizerFlexSort.Add(
            self.tglAwal,
            0,
            wx.ALL | wx.ALIGN_CENTER,
            5,
        )

        self.lblSampaiDgn = wx.StaticText(
            self, wx.ID_ANY, "Sampai dengan", wx.Point(-1, -1), wx.DefaultSize, 0
        )
        self.lblSampaiDgn.Wrap(-1)
        self.tglAkhir = wx.adv.DatePickerCtrl(
            self,
            wx.ID_ANY,
            wx.DefaultDateTime,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.adv.DP_DEFAULT | wx.adv.DP_SHOWCENTURY,
        )
        self.lblKosong1 = wx.StaticText(
            self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0
        )
        self.lblKosong1.Wrap(-1)
        self.lblKosong1.SetMinSize(wx.Size(50, -1))
        rbSinkronChoices = ["Semua", "Belum", "Sinkron"]
        self.rbSinkron = wx.RadioBox(
            self,
            wx.ID_ANY,
            "Status Sinkronisasi",
            wx.DefaultPosition,
            wx.DefaultSize,
            rbSinkronChoices,
            3,
            wx.RA_SPECIFY_COLS,
        )
        self.rbSinkron.SetSelection(0)
        self.lblKosong2 = wx.StaticText(
            self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0
        )
        self.lblKosong2.Wrap(-1)
        self.lblKosong2.SetMinSize(wx.Size(70, -1))
        self.btnCari = wx.Button(
            self, wx.ID_ANY, "Cari", wx.DefaultPosition, wx.DefaultSize, 0
        )
        # Add the widgets to the flex sizer with appropriate alignment
        sizerFlexSort.Add(self.lblSampaiDgn, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        sizerFlexSort.Add(self.tglAkhir, 0, wx.ALL | wx.ALIGN_CENTER, 5)
        sizerFlexSort.Add(self.lblKosong1, 0, wx.ALL, 5)
        sizerFlexSort.Add(self.rbSinkron, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        sizerFlexSort.Add(self.lblKosong2, 0, wx.ALL, 5)
        sizerFlexSort.Add(self.btnCari, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.ALIGN_RIGHT, 5)

        # Do not combine wx.EXPAND with alignment flags for box sizers; use center alignment for the group
        sizerBoxUtama.Add(sizerFlexSort, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        sizerBoxList = wx.BoxSizer(wx.VERTICAL)

        sizerBoxList = wx.BoxSizer(wx.VERTICAL)

        self.lstBendaUji = wx.ListCtrl(
            self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LC_REPORT
        )
        self.lstBendaUji.SetMinSize(wx.Size(-1, 450))

        # idpengujian, tgluji, nodocket, nourutbenda, bujnama, umur,
        # nilaikn,  bebanmpa, kuattekan, beratbenda, tiperetak, sinkron

        self.lstBendaUji.InsertColumn(0, "NO", wx.LIST_FORMAT_RIGHT, wx.LIST_AUTOSIZE)
        self.lstBendaUji.InsertColumn(
            1, "TGL UJI", wx.LIST_FORMAT_CENTRE, wx.LIST_AUTOSIZE
        )
        self.lstBendaUji.InsertColumn(
            2, "DOCKET", wx.LIST_FORMAT_LEFT, wx.LIST_AUTOSIZE
        )
        self.lstBendaUji.InsertColumn(3, "NO BU", wx.LIST_FORMAT_LEFT, wx.LIST_AUTOSIZE)
        self.lstBendaUji.InsertColumn(
            4, "JENIS BU", wx.LIST_FORMAT_LEFT, wx.LIST_AUTOSIZE
        )
        self.lstBendaUji.InsertColumn(5, "UMUR", wx.LIST_FORMAT_LEFT, wx.LIST_AUTOSIZE)
        self.lstBendaUji.InsertColumn(6, "(kN)", wx.LIST_FORMAT_RIGHT, wx.LIST_AUTOSIZE)
        self.lstBendaUji.InsertColumn(
            7, "(Mpa)", wx.LIST_FORMAT_RIGHT, wx.LIST_AUTOSIZE
        )
        self.lstBendaUji.InsertColumn(
            8, "(Kg/cm2)", wx.LIST_FORMAT_RIGHT, wx.LIST_AUTOSIZE
        )
        self.lstBendaUji.InsertColumn(
            9, "BERAT (KG) ", wx.LIST_FORMAT_RIGHT, wx.LIST_AUTOSIZE
        )
        self.lstBendaUji.InsertColumn(
            10, "RETAK", wx.LIST_FORMAT_CENTRE, wx.LIST_AUTOSIZE
        )
        self.lstBendaUji.InsertColumn(
            11, "SINK", wx.LIST_FORMAT_CENTRE, wx.LIST_AUTOSIZE
        )

        sizerBoxList.Add(self.lstBendaUji, 1, wx.ALL | wx.EXPAND, 5)

        sizerBoxUtama.Add(sizerBoxList, 1, wx.EXPAND, 5)

        sizerFlexsum = wx.FlexGridSizer(0, 7, 0, 0)
        sizerFlexsum.SetFlexibleDirection(wx.BOTH)
        sizerFlexsum.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        self.lblJumlahBenda = wx.StaticText(
            self,
            wx.ID_ANY,
            "Jumlah Benda Uji sesuai kondisi diatas",
            wx.DefaultPosition,
            wx.DefaultSize,
            0,
        )
        self.lblJumlahBenda.Wrap(-1)
        self.lblJumlahBenda.SetFont(
            wx.Font(wx.NORMAL_FONT.GetPointSize(), 70, 90, 92, False, wx.EmptyString)
        )
        sizerFlexsum.Add(self.lblJumlahBenda, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        # TextCtrl Berat Benda Uji
        self.txtJumlahBenda = wx.TextCtrl(
            self,
            wx.ID_ANY,
            wx.EmptyString,
            wx.DefaultPosition,
            wx.DefaultSize,
            wx.TE_READONLY,
        )
        self.txtJumlahBenda.SetToolTipString("Berat bersih benda uji")
        sizerFlexsum.Add(self.txtJumlahBenda, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        sizerBoxUtama.Add(sizerFlexsum, 1, wx.ALL | wx.EXPAND, 5)

        self.SetSizer(sizerBoxUtama)
        self.Layout()

        # Connect Events
        self.btnCari.Bind(wx.EVT_BUTTON, self.cariSelect)

    # Virtual event handlers, overide them in your derived class
    def cariSelect(self, paramList):

        paramList = []
        tglAwal = str(self.tglAwal.GetValue())
        tglAwalNya = time.strptime(tglAwal, "%c")
        paramList.append(time.strftime("%Y-%m-%d", tglAwalNya))
        tglAkhir = str(self.tglAkhir.GetValue())
        tglAkhirNya = time.strptime(tglAkhir, "%c")
        paramList.append(time.strftime("%Y-%m-%d", tglAkhirNya))
        s = self.rbSinkron.GetSelection()
        if s == 0:
            paramList.append("%")
        elif s == 1:
            paramList.append("B")
        else:
            paramList.append("S")
        hasilCari = dbctrl.queryGrid(paramList)
        if hasilCari is None:
            hasilCari = []

        self.lstBendaUji.DeleteAllItems()
        for i, r in enumerate(hasilCari, start=1):
            # idpengujian, tgluji, nodocket, nourutbenda, bujnama, umur,
            # nilaikn,  bebanmpa, kuattekan, beratbenda, tiperetak, sinkron
            
            # PERBAIKAN: Ganti sys.maxsize dengan GetItemCount() untuk mendapatkan index yang valid
            # Ini akan menambahkan item di akhir list tanpa menyebabkan error
            index = self.lstBendaUji.InsertStringItem(self.lstBendaUji.GetItemCount(), str(i))
            
            # Format tanggal
            try:
                date1 = datetime.datetime.strptime(str(r[1]), "%Y-%m-%d")
                date2 = datetime.datetime.strftime(date1, "%d/%m/%Y")
                self.lstBendaUji.SetStringItem(index, 1, date2)
            except (ValueError, TypeError) as e:
                logging.warning(f"Error parsing date {r[1]}: {e}")
                self.lstBendaUji.SetStringItem(index, 1, str(r[1]))
            
            # Set data lainnya dengan error handling
            try:
                docketnya = str(r[2]).strip() if r[2] is not None else ""
                self.lstBendaUji.SetStringItem(index, 2, docketnya)
                self.lstBendaUji.SetStringItem(index, 3, str(r[3]) if r[3] is not None else "")
                self.lstBendaUji.SetStringItem(index, 4, str(r[4]) if r[4] is not None else "")
                self.lstBendaUji.SetStringItem(index, 5, str(r[5]) if r[5] is not None else "")
                self.lstBendaUji.SetStringItem(index, 6, str(r[6]) if r[6] is not None else "")
                self.lstBendaUji.SetStringItem(index, 7, str(r[7]) if r[7] is not None else "")
                self.lstBendaUji.SetStringItem(index, 8, str(r[8]) if r[8] is not None else "")
                self.lstBendaUji.SetStringItem(index, 9, str(r[9]) if r[9] is not None else "")
                self.lstBendaUji.SetStringItem(index, 10, str(r[10]) if r[10] is not None else "")
                self.lstBendaUji.SetStringItem(index, 11, str(r[11]) if r[11] is not None else "")
            except (IndexError, TypeError) as e:
                logging.error(f"Error setting list item data: {e}")
                continue

        if hasilCari:
            self.txtJumlahBenda.SetValue(str(len(hasilCari)))
        else:
            self.txtJumlahBenda.SetValue("0")


if __name__ == '__main__':
    app = wx.App(False)
    frame = wx.Frame(None, title='Grid Benda Uji', size=wx.Size(820, 480))
    panel = GridBendaUji(frame)
    frame.Show()
    app.MainLoop()
