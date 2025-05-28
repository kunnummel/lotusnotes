#! python3
from win32com.client import Dispatch
#Author: Jis Joseph Kunnummel - www.kunnummel.com
class UIBack:
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
    if(input("Retrieve backend data for the currently active 'Notes UI'? - (y)/n => ")=='n'):
      return
    UI.reload(False)
    UIBack.db=UI.db
    if(UIBack.db==None):
      return
    if(UIBack.s==None):
      try:
        s.initialize()
        UIBack.s=s
      except Exception as e:
        print(e)
        return
    UIBack.db=s.getdatabase(UIBack.db.server,UIBack.db.filepath) if UIBack.db else None
    if(UI.vw):
      UIBack.vw=UIBack.db.getView(UI.vw.name)
      UIBack.vw.autoupdate=False
      UIBack.vn=UIBack.vw.createviewnav()
      if(UI.dc.count>0 and input(f'Do you also would like to get the {UI.dc.count} documents selected at the UI view? (n)/y >') =='y'):
        UIBack.dc=UIBack.vw.getalldocumentsbykey('\n\n',True)
        UIBack.dcs={}
        Loop.runondocs(UI.dc,lambda x:UIBack.dc.adddocument(UIBack.db.getdocumentbyid(x.noteid)))
        def aa(x):
          UIBack.vn.gotoentry(x)
          return ({x.noteid:list(UIBack.vn.getcurrent().columnvalues)})
        UIBack.dcs=Loop.iterrecords(UIBack.dc,aa)
    try:
      UIBack.doc=UIBack.db.getdocumentbyid(UIBack.vw.caretnoteid)
    except:
      UIBack.doc=UIBack.db.getdocumentbyid(UI.doc.noteid) if UI.doc else None
    if(UIBack.doc and UIBack.vn):
      UIBack.vn.gotoentry(UIBack.doc)
      UIBack.ve=UIBack.vn.getcurrent()
    print(f'''
UIBack.db > {(UIBack.db.server,UIBack.db.filepath,UIBack.db.title,UIBack.db.replicaid) if UIBack.db else 'n/a'}
UIBack.vw > {(UIBack.vw.name,UIBack.vw.aliases,UIBack.vw.entrycount,{k:v for a in [{x.position:x.title+(' (H)' if x.ishidden else '')} for x in UIBack.vw.columns if x.title!='' or x.ishidden] for k,v in a.items()}) if UIBack.vw else 'n/a'}
UIBack.vn/.ve> {(UIBack.ve.getposition('.'),ws.currentview.caretcategory) if UIBack.ve else 'n/a'} 
UIBack.doc> {(UIBack.doc.getitemvalue("form")[0],UIBack.doc.universalid,UIBack.doc.noteid,UIBack.doc.size,{a[0]+1:str(a[1]) for a in enumerate(UIBack.ve.columnvalues)} if UIBack.ve else ('new-doc' if UIBack.doc.isnewnote else 'items-'+str(len(UIBack.doc.items)))) if UIBack.doc else 'n/a'}
UIBack.dc/.dcs > {UIBack.dc.count if UIBack.dc else 'n/a'}
''')

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
  def reload(warn=True):
    if(warn):
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
    UI.vn=UI.vw.createviewnav if UI.vw else None
    UI.ve=(UI.vn.getentry(UI.doc) if UI.doc else UI.vn.getfirst) if UI.vn else None
    if(UI.dc):
      UI.dcs=Loop.iterrecords(UI.dc,lambda x:({x.noteid:list(UI.vn.getentry(x).columnvalues)}))
    print(f'''
UI.db > {(UI.db.server,UI.db.filepath,UI.db.title,UI.db.replicaid) if UI.db else 'n/a'}
UI.vw > {(UI.vw.name,UI.vw.aliases,UI.vw.entrycount,{k:v for a in [{x.position:x.title+(' (H)' if x.ishidden else '')} for x in UI.vw.columns if x.title!='' or x.ishidden] for k,v in a.items()}) if UI.vw else 'n/a'}
UI.vn/.ve> {(UI.ve.getposition('.'),ws.currentview.caretcategory) if UI.ve else 'n/a'} 
UI.doc> {(UI.doc.form[0],UI.doc.universalid,UI.doc.noteid,UI.doc.size,{a[0]+1:str(a[1]) for a in enumerate(UI.ve.columnvalues)} if UI.ve else ('new-doc' if UI.doc.isnewnote else 'items-'+str(len(UI.doc.items)))) if UI.doc else 'n/a'}
UI.dc/.dcs > {UI.dc.count if UI.dc else 'n/a'}
''')
  
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
      return (0,None)
    dc=ws.PickListCollection(3,True,db.server,db.filepath,vname,'Select Document(s)','Pick one or more documents from the view',category)
    return(dc.count,dc)
  def getviewcolumn(colno=1,db=None,vname='',category=None):
    db=UI.db if db==None else db
    vname=UI.vw.name if UI.vw and vname=='' else vname
    if(db==None or vname==''):
      return (0,None)
    return (ws.PickListStrings(3,True,db.server,db.filepath,vname,'Select Value(s)','Pick one or more Values from the view',2,category))
  def files(filter=None,save=False):
    return (ws.SaveFileDialog(False,'Save As',filter,'C:\\')) if save else (ws.OpenFileDialog(True,'Open',filter,'C:\\'))
class Func:
  evaldoc=None
  def getattachments(doc,nodetails=True):
    att=[] if nodetails else {}
    if(doc.hasembedded):
      for x in doc.items:
        if(x.type==1084):
          for y in x.values:
            f=doc.getattachment(y)
            att.append(f) if nodetails else att.update({f.name:(f,f.filesize,f.source)})
    return (att if nodetails else (doc.noteid,att))
  def eval(f,doc=None):
    if(doc==None):
      if(Func.evaldoc==None):
        Func.evaldoc=s.urldatabase.createdocument()
      return s.evaluate(f,Func.evaldoc)
    return s.evaluate(f,doc)
  def export(dc,fields,file='C:/temp/tempfile.xlsx'):
    from openpyxl import Workbook
    wb=Workbook()
    wbs=wb.active
    fields=fields.split(',') if isinstance(fields,str) else fields
    wbs.append(fields)
    def aa(x,w=wbs,f=fields):
      try:
        w.append([str(x.getitemvalue(y)[0]) for y in f])
      except Exception as e:
        print('ERR> ',e)
    Loop.runondocs(dc,aa)
    wb.save(file)
    print ("-- Generated file ",file)
  def dataframe(dcs=None):
    import pandas as pd
    df = pd.DataFrame(dcs if dcs else list(UI.dcs))
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
      
  def docsitemcounts(dc=None,itmname='form'):
    aa={}
    if(dc==None):
      dc=UI.db.search('@all',UI.s.createdatetime(''),0)
    Loop.runondocs(dc,lambda x:aa.update({x.getitemvalue(itmname)[0]:(aa.setdefault(x.getitemvalue(itmname)[0],0)+1)}))
    return (aa)
  def docitemsdict(itms,values=False):
    aa={'text':[],'number':[],'rt':[],'date':[],'name':[],'att':[],'others':[]}
    for x in itms:
      if(x.type in [1280,1282,21]):
        iname='text'
      elif(x.type==768):
        iname='number'
      elif(x.type==1024):
        iname='date'
      elif(x.type==1084):
        iname='att'
      elif(x.type in [1,4,6,7,8,25,1090]):
        iname='rt'
      elif(x.type in [1074,1075,1076]):
        iname='name'
      else:
        iname='others'
      bb=[x.name,x.valuelength,x.type]
      if(values):
        bb.append(x.text)
      aa.get(iname).append(bb)
    return(aa)

  def iterrecords(dc,func,*args,entries=False,details=False):
    if(dc):
      idc=0
      if(entries):
        try:
          e1=dc.getfirstentry()
        except:
          e1=dc.getfirstentry
        while e1:
          idc+=1
          e2=dc.getnextentry(e1)
          yield (idc,e1,func(e1,*args) if func!=None else None) if details else func(e1,*args) if func!=None else None
          e1=e2
      else:
        try:
          d1=dc.getfirstdocument()
        except:
          d1=dc.getfirstdocument
        while d1:
          idc+=1
          d2=dc.getnextdocument(d1)
          yield (idc,d1,func(d1,*args) if func!=None else None) if details else func(d1,*args) if func!=None else None
          d1=d2
  def iternotes(nc,func,*args):
    n1=nc.getfirstnoteid
    while n1!='':
      n2=nc.getnextnoteid(n1)
      yield func(e1,*args)
      n1=n2
  def runondocs(dc,func,*args):
    if(dc and func):
      try:
        idc=0
        d1=dc.getfirstdocument()
      except:
        d1=dc.getfirstdocument
      while d1:
        try:
          d2=dc.getnextdocument(d1)
          idc+=1
          if(Loop.stop(d1.noteid,idc)):
            break
          func(d1,*args)
        except Exception as e:
          print('ERR> ',d1.noteid,idc,e)	
          if(input('Do you wish to break it? (y)/n >')!='n'):
            break
        d1=d2
  def runonnotes(nc,func,*args):
    if(nc and func):
      n1=nc.getfirstnoteid
      while n1:
        try:
          n2=nc.getnextnoteid(n1)
          if(Loop.stop(n1)):
            break
          func(n1,*args)
        except Exception as e:
          print('ERR> ',n1,e)	
          if(input('Do you wish to break it? (y)/n >')!='n'):
            break
        n1=n2
  def runonentries(ec,func,*args):
    if(ec and func):
      try:
        idc=0
        e1=ec.getfirstentry()
      except:
        e1=ec.getfirstentry
      while e1:
        try:
          e2=ec.getnextentry(e1)
          idc+=1
          if(Loop.stop(e1.noteid,idc)):
            break
          func(e1,*args)
        except Exception as e:
          print('ERR> ',e1.noteid,idc,e)
        e1=e2
  def stop(*args,key='ctrl+c'):
    try:
      import keyboard
      if(keyboard.is_pressed(key)):
        if(input(f'>> Do you wish to interrupt? {args} (n)/Y > ')=='Y'):
          print('-- Stopped at user request at ',args)
          return True
    except Exception as e:
      print('ERR> ',e)
      return True
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
  def comprops(obj=[],prop=['name']):
    if(not isinstance(obj,(list,tuple))):
      obj=[obj]
    if(not isinstance(prop,(list,tuple))):
      prop=[prop]
    _obj=[]
    for x in obj:
      for y in prop:
        try:
          _obj.append(x.__getattr__(y))
        except Exception as e:
          print('ERR> ',x,y,e)
      if(Loop.stop(x,y)):
        return _obj
    return _obj
try:
  ws=Dispatch('Notes.NotesUIWorkspace')
  UI.s=Dispatch('Notes.NotesSession')
  s=Dispatch('Lotus.NotesSession')
except Exception as e:
  print('ERR> ',e)
finally:
  print(f'Global>> s ws UI UIBack Loop Func')
  del Dispatch
UI.reload()
