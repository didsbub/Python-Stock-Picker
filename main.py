import wx
import wx.lib.platebtn as platebtn
import wx.lib.dialogs
import thread
import pleco
import sys, os
import shutil
import tempfile
import time

class GroupBox(wx.StaticBox):
    def __init__(self, parent, boxcaption):
        wx.StaticBox.__init__(self, parent, wx.ID_ANY, boxcaption)
        self.sizer = wx.StaticBoxSizer(self, wx.HORIZONTAL)

        checkboxcaptionlist = ["Company", "Financials", "Prices", "Extra Info"]
        self.checkboxes={}
        for caption in checkboxcaptionlist:
            self.checkboxes[caption] = wx.CheckBox(parent, wx.ID_ANY, caption)
            self.sizer.Add(self.checkboxes[caption], 1, wx.EXPAND)

        self.selectallbutton = platebtn.PlateButton(parent, wx.ID_ANY,"  SELECT ALL  ",
                style=platebtn.PB_STYLE_DEFAULT|platebtn.PB_STYLE_TOGGLE) 
        self.sizer.Add(self.selectallbutton,0, wx.CENTER)

        self.selectallbutton.Bind(wx.EVT_TOGGLEBUTTON, self.SelectAll)

    def SelectAll(self, evt):
        button = evt.GetEventObject()
        ispressed = button.IsPressed()

        if ispressed:
        #If the button is pressed now to select all checkboxes
            for caption, checkbox in self.checkboxes.items():
                checkbox.SetValue(True)
            button.SetLabel("SELECT NONE")
        else:
        #If the button is pressed now to deselect all checkboxes
            for caption, checkbox in self.checkboxes.items():
                checkbox.SetValue(False)
            button.SetLabel("  SELECT ALL  ")

    def GetSelection(self):
        selectionlist = []
        for caption, checkbox in self.checkboxes.items():
            if checkbox.GetValue() == True:
                selectionlist.append(caption)
        return selectionlist

    def Disable(self):
        self.Enable(False)

    def Enable(self, enable=True):
        wx.StaticBox.Enable(self, enable)
        if enable == False: #Disable:
            for checkbox in self.checkboxes.values()+[self.selectallbutton]:
                checkbox.Disable()
        else:
            for checkbox in self.checkboxes.values()+[self.selectallbutton]:
                checkbox.Enable()

class PlecoFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, size = (600, 350), 
                title="Pleco Stock Data Scraper")

        self.panel = PlecoPanel(self)

        self.statusbar = wx.StatusBar(self)
        self.SetStatusBar(self.statusbar)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def OnClose(self, evt):
        if evt.CanVeto() and (self.panel.updating or self.panel.exporting):
            dlg = wx.MessageDialog(self, 'There is an ongoning update or export process. Are you sure?',
                               'Close Application',
                               wx.YES_NO | wx.NO_DEFAULT | wx.ICON_EXCLAMATION
                               )
            if dlg.ShowModal() == wx.ID_YES:
                evt.Skip()
            dlg.Destroy()
        else:
            evt.Skip()



class Redirect:
    def __init__(self, writefunction):
        self.writefunction = writefunction 
    def write(self, string):
        if string == "\n": return
        self.writefunction(string)
    def flush(self):
        pass

class PlecoPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.SetSizer(sizer)
        self.log = []
        self.controls = set()

        self.parent = parent
        self.updating = False
        self.exporting = False

        sys.stdout = Redirect(self.StandardOutput)
        sys.stderr = Redirect(self.StandardErr)

        self.stockexchangelist = ["TSX", "NYSE", "NASDAQ", "HKG"]
        self.stockboxes = {}
        for stock in self.stockexchangelist: 
            self.stockboxes[stock] = GroupBox(self, stock)
            sizer.Add(self.stockboxes[stock].sizer, 1, wx.EXPAND)
            self.controls.add( self.stockboxes[stock] )


        horsizer1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(horsizer1, 1, wx.CENTER)

        self.updatebutton = wx.Button(self, label="UPDATE")
        self.showlogbutton = wx.Button(self, label="SHOW LOG")
        self.emptybutton = wx.Button(self, label="EMPTY CACHE")
        self.exportbutton = wx.Button(self, label="EXPORT TO EXCEL")

        self.controls.add( self.updatebutton )
        self.controls.add( self.showlogbutton )
        self.controls.add( self.emptybutton )
        self.controls.add( self.exportbutton )


        horsizer1.Add(self.updatebutton, 1, wx.CENTER)
        horsizer1.Add(self.showlogbutton, 1, wx.CENTER)
        horsizer1.Add(self.emptybutton, 1, wx.CENTER)
        horsizer1.Add(self.exportbutton, 1, wx.CENTER)


        self.Bind(wx.EVT_BUTTON, self.UpdateButtonClicked, self.updatebutton)
        self.Bind(wx.EVT_BUTTON, self.ShowLogClicked, self.showlogbutton)
        self.Bind(wx.EVT_BUTTON, self.EmptyCache, self.emptybutton)
        self.Bind(wx.EVT_BUTTON, self.ExportToExcel, self.exportbutton)

    def EmptyCache(self, evt):
        dlg = wx.MessageDialog(self, 
        u'Saved web pages will be deleted. Are you sure?', 
        u'Confirm Delete Cache Contents',
        wx.YES_NO | wx.ICON_QUESTION | wx.NO_DEFAULT)
        yesno = dlg.ShowModal()
        dlg.Destroy()
        if yesno == wx.ID_YES:
            pleco.PageCache().EmptyCache()
            dlg = wx.MessageDialog(self, 'Cache has been emptied successfully!',
                               'Success',
                               wx.OK | wx.ICON_INFORMATION
                               )
            dlg.ShowModal()
            dlg.Destroy()


    def ExportToExcel(self, evt):
        def GenerateSheet(self, excelfile, arguments):
            pleco.SetXlsFilename(excelfile)
            pleco.run(arguments)
            SaveSheet(excelfile)
            #export completed successfully:
            self.exporting = False
            self.EnableAllControls()
            self.exportbutton.SetLabel("EXPORT TO EXCEL")
        
        def SaveSheet(excelfile):
            wildcard = "Excel File (*.xls)|*.xls|"\
                       "All files (*.*)|*.*"
            dlg = wx.FileDialog( self, message="Save excel file as ...",
                    defaultDir=os.getcwd(), defaultFile="pleco.xls", wildcard=wildcard,
                    style=wx.SAVE)
            if dlg.ShowModal() == wx.ID_OK:
                savefilename = dlg.GetPath()
                try: #check if the selected file already exists
                    os.stat(savefilename)
                    fileexists=True
                except: 
                    fileexists=False
                finally:
                    if fileexists:
                        dlg = wx.MessageDialog(self, 
                        u'File you have selected already exists. Do you want to replace it?', 
                        u'Confirm Save',
                        wx.YES_NO | wx.ICON_QUESTION | wx.NO_DEFAULT)
                        yesno = dlg.ShowModal()
                        dlg.Destroy()
                        if yesno != wx.ID_YES:
                            return SaveSheet(excelfile)
                        else:
                            shutil.copyfile(excelfile, savefilename)
                            ShowSuccess()
                    else:
                        shutil.copyfile(excelfile, savefilename)
                        ShowSuccess()

        def ShowSuccess():
            dlg = wx.MessageDialog(self, 'Export to excel completed succesfully!',
                               'Success',
                               wx.OK | wx.ICON_INFORMATION
                               )
            dlg.ShowModal()
            dlg.Destroy()


        def ExportExcel(self):
            dlg = wx.MultiChoiceDialog( self, 
                                       "Select stock exchanges to be exported",
                                       "Export To Excel", self.stockexchangelist)
            dlg.SetSelections(range(0,len(self.stockexchangelist)))#by default select all

            if (dlg.ShowModal() == wx.ID_OK):
                self.exporting = True
                self.exportbutton.SetLabel("CANCEL EXPORT")
                self.DisableAllControlsExcept((self.exportbutton, self.showlogbutton))
                selections = dlg.GetSelections()
                mylist=[]
                for i in selections:
                    mylist.append(self.stockexchangelist[i])

                excelfile = tempfile.NamedTemporaryFile(suffix=".xls").name
                arguments = [None,"--excelexport@"+",".join(mylist)]
                thread.start_new_thread(GenerateSheet,(self, excelfile,arguments))
            else:
                pass

            dlg.Destroy()


        if not self.exporting:
            ExportExcel(self)
        else:
            dlg = wx.MessageDialog(self, 'Are you sure? If yes, the application will close.',
                               'Cancel Export?',
                               wx.YES_NO | wx.NO_DEFAULT | wx.ICON_EXCLAMATION
                               )
            if dlg.ShowModal() == wx.ID_YES:
                print >> sys.stderr, "Export terminated upon user request."
                self.exporting = False
                self.exportbutton.SetLabel("EXPORT TO EXCEL")
                self.EnableAllControls()
                self.parent.Close()
            else:
                dlg.Destroy()
        
    def StandardOutput(self, text):
        self.log.append(text)
        self.SetStatusText(text)

    def StandardErr(self, text):
        self.StandardOutput(text)

    def SetStatusText(self, text):
        self.parent.SetStatusText(text,0)

    def GetCommandlineArguments(self):
        argumentlist = [None]
        for stockexchange in self.stockexchangelist:
            selection = self.stockboxes[stockexchange].GetSelection()
            if "Company" in selection:
                argumentlist.append("--companies@%s"%stockexchange)
            if "Financials" in selection:
                argumentlist.append("--financials@%s"%stockexchange)
            if "Prices" in selection:
                argumentlist.append("--prices@%s"%stockexchange)
            if "Extra Info" in selection:
                argumentlist.append("--extra@%s"%stockexchange)
        return argumentlist

    def DisableAllControlsExcept(self,excludedcontrols):
        controlstobedisabled = self.controls - set(excludedcontrols)
        for control in controlstobedisabled:
            control.Disable()

    def EnableAllControls(self):
        for control in self.controls:
            control.Enable()

    def UpdateButtonClicked(self, evt):
        def Update(self, updatebutton):
            self.updating = True
            try:
                pleco.run(self.GetCommandlineArguments())
                #Updating Completed Successfully:
                self.updating = False
            except StandardError, e:
                #Update failed
                self.updating = False
                print >> sys.stderr, e
                print >> sys.stderr, 'An error occurred! Update did not complete successfully.'
                dlg = wx.MessageDialog(self, 'An error occurred! Update did not'\
                    ' complete successfully and application will close.'\
                    ' If problem persists  please consult to programmer'\
                    '  with application id number = 1127 via skype account: emre062011.\n\nLast log entries:\n%s'%"\n".join(self.log[-10:]),
                                  'Failure',
                                   wx.OK | wx.ICON_INFORMATION
                                   )
                if dlg.ShowModal() == wx.ID_OK:
                    self.parent.Close()
                dlg.Destroy()
                return
            updatebutton.SetLabel("Update")
            self.EnableAllControls()
            print >> sys.stdout, 'Update completed succesfully!'
            dlg = wx.MessageDialog(self, 'Update completed succesfully!',
                               'Success',
                               wx.OK | wx.ICON_INFORMATION
                               )
            dlg.ShowModal()
            dlg.Destroy()
        
        if not self.updating:
            updatebutton = evt.GetEventObject()
            thread.start_new_thread(Update,(self, updatebutton))
            updatebutton.SetLabel("CANCEL UPDATE")
            self.DisableAllControlsExcept((self.updatebutton, self.showlogbutton))
        else:
            dlg = wx.MessageDialog(self, 'Are you sure? If yes, the application will close.',
                               'Cancel Update?',
                               wx.YES_NO | wx.NO_DEFAULT | wx.ICON_EXCLAMATION
                               )
            if dlg.ShowModal() == wx.ID_YES:
                print >> sys.stderr, "Update terminated upon user request."
                self.updating = False
                evt.GetEventObject().SetLabel("Update")
                self.EnableAllControls()
                self.parent.Close()
            else:
                dlg.Destroy()

    def ShowLogClicked(self, evt):
        msg = "\n".join(self.log)
        dlg = wx.lib.dialogs.ScrolledMessageDialog(self, msg, "Pleco output log")
        dlg.ShowModal()
        dlg.Destroy()

if __name__ == "__main__":
    app = wx.App(False)
    frame = PlecoFrame(None)
    frame.Show()
    app.MainLoop()

    

