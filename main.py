import sys

from PyQt5 import QtWidgets,QtGui
import glob
import  pandas as pd
import numpy as np

class Pencere (QtWidgets.QWidget):
    def __init__(self):

        super().__init__()


        self.init_ui()
    def init_ui(self):
        self.aciklama2 = QtWidgets.QLabel("")
        self.table=QtWidgets.QTableWidget() 
        self.gorsel = QtWidgets.QLabel("Logo")
        self.butoncalistir = QtWidgets.QPushButton("RUN")
        self.butoncalistir.setIcon(QtGui.QIcon("calistir.png"))
        self.butonkaydet = QtWidgets.QPushButton("save")
        self.butonkaydet.setIcon(QtGui.QIcon("kaydet.png"))
        self.aciklama=QtWidgets.QLabel("Select the relevant unit:")
        self.gorsel.setPixmap(QtGui.QPixmap("logo-min.png"))
        self.combo_box =QtWidgets.QComboBox()
        bolumler=[]
        for name in glob.glob("/*.xlsx"):
             sayac=0
             z=name[::-1]
             for j in name:
                 if j==("\\"):
                     indis=sayac
                 sayac+=1
             bolumler.append(name[indis+1:-5])
                
      
        self.combo_box.addItems(bolumler)

        self.butonyazdir=QtWidgets.QPushButton("Yazdır")





        h_box = QtWidgets.QHBoxLayout()
        h_box.addWidget(self.aciklama)
        h_box.addWidget(self.combo_box)
        h_box.addWidget(self.butoncalistir)
        h_box.addWidget(self.butonkaydet)
        h_box.addWidget(self.gorsel)

        h_box2=QtWidgets.QHBoxLayout()
        
        h_box2.addWidget(self.aciklama2)
        h_box2.addStretch()
        h_box2.addWidget(self.butonyazdir)

        v_box=QtWidgets.QVBoxLayout()

        v_box.addLayout(h_box)
        v_box.addWidget(self.table)
        v_box.addStretch()
        v_box.addLayout(h_box2)
        self.setLayout(v_box)
        self.setWindowTitle("Nursce Scheduling")
        self.setGeometry(600,150,700,700)
        
        self.butoncalistir.clicked.connect(self.click)
        self.butonkaydet.clicked.connect(self.click)
        self.butonyazdir.clicked.connect(self.click)
        self.show()
    def click (self):   
        sender = self.sender()
        
        if sender.text() == "Run":
            x=self.combo_box.currentText()
            x= "%s.xlsx"%(x)
            
            
            yi=pd.read_excel(x,sheet_name="yillikizin")
            rapor=pd.read_excel(x,sheet_name="rapor")
            cmmgu=pd.read_excel(x,sheet_name="calismamagunduz")
            cmmge=pd.read_excel(x,sheet_name="calismamagece")
            
            thgu=pd.read_excel(x,sheet_name="tophemgun")
            thge=pd.read_excel(x,sheet_name="tophemgece")
            
            
            yi=np.array(yi)
            rapor=np.array(rapor)
            cmmgu=np.array(cmmgu)
            cmmge=np.array(cmmge)
            thgu=np.array(thgu)
            thge=np.array(thge)
            
            
            hems=len(cmmgu)
            gun=len(yi[0][:])
            
            geat=np.zeros([hems,gun])
            guat=np.zeros([hems,gun])
            ideca=np.zeros(hems)                              
            mesa=np.zeros(hems)
            pazar=np.zeros([hems,gun])
            hson=[6,13,20,27]
            hsonc=np.array([0,0,0,0])
            
            for i in range(hems):
                toplam=0
                for j in hson:
                   if (yi[i][j]==0 and yi[i][j-1]==0 and yi[i][j-2]==0) or (rapor[i][j]==0 and rapor[i][j-1]==0 and rapor[i][j-2]==0):
                       toplam+=1
                    
                if toplam>0:
                    r1=np.argmin(hsonc)
                    pazar[i][hson[r1]]=1
                    pazar[i][hson[r1]-1]=1
                    pazar[i][hson[r1]-2]=1
                    hsonc[r1]+=1                   
            
            
            
            for i in range(hems):
                izgu=sum(rapor[i][:])+sum(yi[i][:])
                ideca[i]=(28-izgu)*(160/28)
                
                
                
            def varmi(a,b):
                for  i in range(len(a)):
                    d=0
                    for j in range(len(b)):
                        if a[i]==b[j]:
                            d=1
                            break
                    if d==0:
                        b.append(a[i])
                return b
            
            def mesahesap(a,m,ideal):
                meuz=[]
                for i in a:
                    kalanmesai= abs((ideal[i]-m[i]))/ideal[i]
                    meuz.append(kalanmesai)
                meuz=np.array(meuz)
                return meuz
            
            def guncelleme(gunn,d1,d2,d3,d4,athgu,athge,sayac1,mesa,ideca,thgu,thge,guat,geat):
            
                if sayac1==1:
                    if d1==1 or d3==1:
                      athgu2=[]
                      for j in range(hems):
                        if rapor[j][gunn]!=1 and yi[j][gunn]!=1 and pazar[j][gunn]!=1:                    # j.hemşire i.gün raporluysa i .güne atanamaz.
                             if gunn>0 and geat[j][gunn-1]!=1 :
                                     athgu2.append(j)  
                                     
                             elif gunn==0:
                                     athgu2.append(j)
                    
                    if d2==1 or d4==1:
                      athge2=[]
                      for j in range(hems):
                        if rapor[j][gunn]!=1 and yi[j][gunn]!=1 and pazar[j][gunn]!=1:                    # j.hemşire i.gün raporluysa i .güne atanamaz.
                             if gunn>0 and geat[j][gunn-1]!=1 :
                                     athge2.append(j)  
                                     
                             elif gunn==0:
                                     athge2.append(j)
                    if d1==1 or d3==1:
                        athgu=varmi(athgu2,athgu)
                    if d2==1 or d4==1:
                        athge=varmi(athge2,athge) 
                        
                elif sayac1==2:                 
                    if d1==1 or d3==1:
                      athgu2=[]
                      for j in range(hems):
                        if rapor[j][gunn]!=1:                                                              # j.hemşire i.gün raporluysa i .güne atanamaz.
                             if gunn>0 and geat[j][gunn-1]!=1 :
                                     athgu2.append(j)  
                                     
                             elif gunn==0:
                                     athgu2.append(j)
                    
                    if d2==1 or d4==1:
                      athge2=[]
                      for j in range(hems):
                        if rapor[j][gunn]!=1 :                                                             # j.hemşire i.gün raporluysa i .güne atanamaz.
                             if gunn>0 and geat[j][gunn-1]!=1 :
                                     athge2.append(j)  
                                     
                             elif gunn==0:
                                     athge2.append(j)
                    if d1==1 or d3==1:
                        athgu=varmi(athgu2,athgu)
                    if d2==1 or d4==1:
                        athge=varmi(athge2,athge)   
                              
                d1=0
                d2=0
                d3=0
                d4=0              
                return d1,d2,d3,d4,athgu,athge,sayac1
                
                
                
                
            def atama(gunn,athgu,athge,mesa,ideca,thgu,thge,guat,geat,hems) :      
                                                     
                sayac3=0
              
                while True:
                    meuzgun=mesahesap(athgu,mesa,ideca)
                    meuzgunindex=np.argsort(-meuzgun)
                    meuzgece=mesahesap(athge,mesa,ideca)
                    meuzgeceindex=np.argsort(-meuzgece)
                    for i in range(hems):
                        guat[i][gunn]=0
                        geat[i][gunn]=0
                    
                    a=len(athgu)
                    b=len(athge)
            
                    d1=0
                    d2=0
                    d3=0
                    d4=0
                    
                    if a<thgu[0][gunn]:
                        d1=1
                    if b<thge[0][gunn]:
                        d2=1
                    
                    if a>=thgu[0][gunn] and b>=thge[0][gunn]:
                        
                        if b>=a:
                            for i in range(thgu[0][gunn]):
                                guat[athgu[meuzgunindex[i]]][gunn]=1
                                mesa[athgu[meuzgunindex[i]]]= mesa[athgu[meuzgunindex[i]]]+12
                            sayac1=0
                            sayac2=0
                            while sayac1<b and sayac2<thge[0][gunn]:
                               if  guat[athge[meuzgeceindex[sayac1]]][gunn]!=1 :
                                   geat[athge[meuzgeceindex[sayac1]]][gunn]=1
                                   mesa[athge[meuzgeceindex[sayac1]]]= mesa[athge[meuzgeceindex[sayac1]]]+12
                                   sayac2+=1
                               sayac1+=1
                            if sayac2<thge[0][gunn]:
                                d3=1
                        elif a>b:
                            for i in range(thge[0][gunn]):
                                geat[athge[meuzgeceindex[i]]][gunn]=1
                                mesa[athge[meuzgeceindex[i]]]= mesa[athge[meuzgeceindex[i]]]+12
                            sayac1=0
                            sayac2=0
                            while sayac1<a and sayac2<thgu[0][gunn] :
                               if  geat[athgu[meuzgunindex[sayac1]]][gunn]!=1 :
                                   guat[athgu[meuzgunindex[sayac1]]][gunn]=1
                                   mesa[athgu[meuzgunindex[sayac1]]]= mesa[athgu[meuzgunindex[sayac1]]]+12
                                   sayac2+=1
                               sayac1+=1
                            if sayac2<thgu[0][gunn]:
                                d4=1
                    if d1==0 and d2==0 and d3==0 and d4==0:
                        break
                    else:
                        sayac3+=1
                        d1,d2,d3,d4,athgu,athge,sayac3=guncelleme(gunn,d1,d2,d3,d4,athgu,athge,sayac3,mesa,ideca,thgu,thge,guat,geat)                                                    
                            
                return   guat, geat, mesa 
            
            
            for j in range(gun):
            
               
                    athgu=[]
                    athge=[]
                    for i in range(hems):
                        if rapor[i][j]!=1 and yi[i][j]!=1 and pazar[i][j]!=1 :                       # j.hemşire i.gün raporluysa i .güne atanamaz.
                             if j>0 and geat[i][j-1]!=1 :
                                 if cmmgu[i][j]!=1:
                                     athgu.append(i)
                                 if cmmge[i][j]!=1:
                                     athge.append(i)   
                                     
                             elif j==0:
                                if cmmgu[i][j]!=1:
                                     athgu.append(i)
                                if cmmge[i][j]!=1:
                                     athge.append(i)  
                                     
                    guat, geat ,mesa=atama(j,athgu,athge,mesa,ideca,thgu,thge,guat,geat,hems)
                    
            self.aciklama2.setText("Çizelgeleme yapıldı")
            self.guat=guat
            self.geat=geat
            self.hems=hems
            self.gun=gun
            self.mesa=mesa
            self.yi=yi
            self.rapor=rapor
        elif sender.text()== "Kaydet"  :           
            columns=["Çalışma Saati"]
            
            for i in range(self.gun):
                columns.append("%s .Gün "%(i+1))
            
            
            index=[]
            
            for i in range(self.hems):
                index.append("%s .Hemşire "%(i+1))
            
            data=[([0]*(self.gun+1)) for i in range(self.hems)]
            
            for i in range(self.hems):
                for j in range(self.gun):
                    
                    if self.geat[i][j]==1:
                        data[i][j+1]=2
                    if self.guat[i][j]==1:
                        data[i][j+1]=1
                    if self.yi[i][j]==1 and   self.guat[i][j]==0 and self.geat[i][j]==0: 
                        data[i][j+1]="S"
                    if self.rapor[i][j]==1:
                         data[i][j+1]="R"
                data[i][0]=self.mesa[i]         
            cikti = pd.DataFrame(data=data,index=index,columns=columns)
           
            a=self.combo_box.currentText()
            a= "%s sonuc.xlsx"%(a)
            cikti.to_excel(a)
            self.aciklama2.setText(" Kayıt yapıldı")    
            self.data=data
            a=444545
        elif sender.text()=="Yazdır":
          self.table.setRowCount(self.hems)
          self.table.setColumnCount(self.gun+1)
          for i in range(self.hems):
              for j in range(self.gun+1):
                  a=str(self.data[i][j])
                  self.table.setItem(i,j,QtWidgets.QTableWidgetItem(a))



app = QtWidgets.QApplication(sys.argv)
pencere = Pencere()
sys.exit(app.exec_())