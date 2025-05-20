#! python3
from win32com.client import Dispatch
#Author: Jis Joseph Kunnummel
class UI:
  s=None
  db=None
  mdb=None
  mvw=None
  mnav=None
  vw=None
  vn=None
  ve=None
  doc=None
  dc=None
  dcs=None
  def reload():
    if(input("Retrieve currently active 'Notes UI' data? - (y)/n => ")=='n'):
      return
    while True:
      UI.db=ws.currentdatabase
      if(UI.db or input("No active Database objects found. Try again (y)/n => ")=='n'):
          break
    if(UI.db==None):
      return
    UI.vw=ws.currentview
    UI.dc=UI.vw.documents if UI.vw else None
    UI.db=UI.db.database if UI.db else None
    try:
      UI.doc=UI.db.getdocumentbyid(UI.vw.caretnoteid)
    except:
      UI.doc=ws.currentdocument
      UI.doc=UI.doc.document if UI.doc else None
    if(UI.vw):
      UI.vw=UI.vw.view
      UI.vw.autoupdate=False
      UI.vn=UI.vw.createviewnav
    UI.ve=UI.vn.getentry(UI.doc) if UI.vn and UI.doc else None
    if(UI.dc):
      UI.dcs={}
      Loop.runondocs(UI.dc,lambda x:UI.dcs.update({x.noteid:list(UI.vn.getentry(x).columnvalues)}))
    print(f'''UI.db > {(UI.db.server,UI.db.filepath,UI.db.title,UI.db.replicaid) if UI.db else 'n/a'}\nUI.vw > {((UI.vw.name,)+(UI.vw.aliases,UI.vw.entrycount)+({k:v for a in [{x.position:x.title} for x in UI.vw.columns] for k,v in a.items()},)) if UI.vw else 'n/a'}\nUI.doc> {(UI.doc.form[0],UI.doc.universalid,UI.doc.noteid,UI.doc.size,{a[0]+1:a[1] for a in enumerate(UI.ve.columnvalues)} if UI.ve else len(UI.doc.items)) if UI.doc else 'n/a'}\nUI.dc/.dcs > {UI.dc.count if UI.dc else '0'}''')
  
  def checkmail(cnt=5):
    if(UI.db and (not UI.mdb)):
      UI.mdb=UI.s.getdatabase('','')
      if(not UI.mdb.isopen):
        UI.mdb.openmail
    if (not UI.mdb):
      return
    if(not UI.mnav):
      UI.mvw=UI.mdb.getview('($Inbox)')
      if(UI.mvw==None):
        return
      UI.mnav=UI.mvw.createviewnav
    ve=UI.mnav.getlast
    for x in range(cnt):
      if(ve):
        doc=ve.document
        print(doc.posteddate[0],' - ', doc.getitemvalue('from')[0],'\n',doc.subject[0],'\n',doc.getitemvalue('$abstract')[0],'\n','-'*100,'\n')
        ve=UI.mnav.getprev(ve)
  def getviewdocs(db=None,vname='',category=None):
    db=db if db else UI.db
    vname=UI.vw.name if UI.vw and vname=='' else vname
    if(db==None or vname==''):
      return
    dc=ws.PickListCollection(3,True,db.server,db.filepath,vname,'Select Document(s)','Pick one or more documents from the view',category)
    return(dc.count,dc)
  def getviewcolumn(colno=1,db=None,vname='',category=None):
    db=UI.db if db==None else db
    vname=UI.vw.name if UI.vw and vname=='' else vname
    if(db==None or vname==''):
      return
    return (ws.PickListStrings(3,True,db.server,db.filepath,vname,'Select Value(s)','Pick one or more Values from the view',2,category))
  def files(filter=None,save=False):
    return (ws.SaveFileDialog(False,'Save As',filter,'C:\\')) if save else (ws.OpenFileDialog(True,'Open',filter,'C:\\'))
class Func:
  def getattachments(doc):
    att={}
    if(doc.hasembedded):
      for x in doc.items:
        if(x.type==1 and x.embeddedobjects!=None):
          for f in x.embeddedobjects:
            if(f.type==1454):
              att.update({f.name:(f.filesize,f)})
    return (att)
  def evaluate(f,doc=None):
    if(doc==None):
      doc=s.getdatabase('','perweb.nsf').createdocument()
    return s.evaluate(f,doc)
  def export(dc,fields,file=r'C:\temp\tempfile.xlsx'):
    from openpyxl import Workbook
    wb=Workbook()
    wbs=wb.active
    fields=fields.split(',') if isinstance(fields,str) else fields
    wbs.append(fields)
    def aa(x,w=wbs,f=fields):
      try:
        w.append([str(x.getitemvalue(y)[0]) for y in f])
      except Exception as e:
        print(e)
    Loop.runondocs(dc,aa)
    wb.save(file)
    print ("Generated file ",file)
  def dataframe(dcs=None):
    import pandas as pd
    df = pd.DataFrame(dcs if dcs else UI.dcs)
    return df.transpose()
class Loop:
  def dbdirectory(server='',type=1247):
    dbr=UI.s.getdbdirectory(server)
    db=dbr.getfirstdatabase(type)
    dbrdict={}
    while db:
      dbrdict.update({db.filepath:(db.title,db)})
      db=dbr.getnextdatabase
    return dbrdict
  
  def docitemsdict(doc):
    aa={}
    for x in itms:
      aa.update({x.name:(list(x.values) if len(x.values)>1 else (x.values[0] if len(x.values)==1 else '')) if isinstance(x.values,tuple ) else x.values})
    return(aa)

  def iternotes(dc,entries=False):
    if(dc):
      idc=0
      if(entries):
        try:
          e1=dc.getfirstentry()
        except:
          e1=dc.getfirstentry
        while e1:
          idc+=1
          yield (idc,e1)
          e1=dc.getnextentry(e1)
      else:
        try:
          d1=dc.getfirstdocument()
        except:
          d1=dc.getfirstdocument
        while d1:
          idc+=1
          yield (idc,d1)
          d1=dc.getnextdocument(d1)
  def runondocs(dc,func):
    if(dc):
      try:
        idc=0
        d1=dc.getfirstdocument()
      except:
        d1=dc.getfirstdocument
      while d1:
        try:
          idc+=1
          func(d1)
          if(Loop.stop(d1.noteid,idc)):
            break
        except Exception as e:
          print(d1.noteid,idc,e)
          continue
        d1=dc.getnextdocument(d1)
  def runonentries(ec,func):
    if(ec):
      try:
        idc=0
        e1=ec.getfirstentry()
      except:
        e1=ec.getfirstentry
      while e1:
        try:
          idc+=1
          func(e1)
          if(Loop.stop(e1.noteid,idc)):
            break
        except Exception as e:
          print(e1.noteid,idc,e)
          continue
        e1=ec.getnextentry(e1)
  def stop(*args,key='esc'):
    try:
      import keyboard
      if(keyboard.is_pressed(key)):
        if(input(f'Do you wish to interrupt? {args} (n)/Y >')=='Y'):
          print('Stopped at user request at ',args)
          return True
    except:
      pass
  def getdocids(dc,unid=False,entries=False):
    selids=[]
    if(entries):
      Loop.runonentries(dc,lambda e1:selids.append(e1.document.universalid if unid else e1.document.noteid))
    else:
      Loop.runondocs(dc,lambda d1:selids.append(d1.universalid if unid else d1.noteid))
    return(selids)

  def getiddocs(selids=[],db=UI.db):
    seldocs={}
    if (db==None):
      db=UI.db
    for id in selids:
      seldocs.update({id:db.getdocumentbyunid(id) if len(id)==32 else db.getdocumentbyid(id)})
    return(seldocs)

try:
  ws=Dispatch('Notes.NotesUIWorkspace')
  UI.s=Dispatch('Notes.NotesSession')
  s=Dispatch('Lotus.NotesSession')
except Exception as e:
  print(e)
finally:
  print(f'Global>> s ws UI Loop Func')
  del Dispatch
UI.reload()
