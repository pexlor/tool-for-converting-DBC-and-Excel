#include "exceltodbc_sence.h"
#include <QPainter>
#include <QPushButton>
#include <QFileDialog>
#include <QDebug>
#include "excel.h"
#include <QDir>
#include <QMessageBox>
#include <QTextEdit>
#include <QApplication>
Exceltodbc_sence::Exceltodbc_sence(QWidget *parent) : QMainWindow(parent)
{
    /*场景布置*/
    this->setFixedSize(500,500);
    this->setWindowIcon(QIcon(":/res/Set.bmp"));
    this->setWindowTitle("Dbc To Excel");

    excel=new Excel();

    QPushButton *ChooseBtn = new QPushButton("选择Excel文件",this);
    ChooseBtn->setFixedSize(150,80);
    ChooseBtn->move(30,this->height()/5);

    QPushButton *BeginBtn = new QPushButton("输出为Dbc",this);
    BeginBtn->setFixedSize(150,80);
    BeginBtn->move(this->width()-180,this->height()/5);

    textEdit=new QTextEdit(this);
    textEdit->setFixedSize(300,200);
    textEdit->move(100,200);

    connect(ChooseBtn,&QPushButton::clicked,[=](){
        //弹出文件选择框
        QString curPath=QDir::currentPath();//获取系统当前目录
        //获取应用程序的路径
        QString dlgTitle="选择一个dbc文件"; //对话框标题
        QString filter="*.xlsx"; //文件过滤器
        QString aFileName=QFileDialog::getOpenFileName(this,dlgTitle,curPath,filter);
        if (!aFileName.isEmpty()){
           QString ExclePath=aFileName;
           //qDebug()<<ExclePath;
           excel->readExcel(ExclePath,"Sheet1");
        }
        textEdit->append("读取成功！");
    });
    connect(BeginBtn,&QPushButton::clicked,[=](){
           Dbc_Analyse();
           SaveDbc();
           textEdit->append("输出成功！");
    });
}



 void Exceltodbc_sence:: Dbc_Analyse()
 {
     int clomn=0;
     QVector<QVector<QString>> vecDatas=excel->vecDatas;

     /*根据G3项目的EXCEL*/
     //QVector<int> ReadMessageOrder={0,1,2,3,4,5,6};
     //QVector<int> ReadSingalOrder={7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27};

     /*
        下面的步骤与SetInfo函数中的步骤对应
        1.报文名称
        2.报文类型
        3.报文ID
        4.参数组编号
        5.报文发送类型
        6.报文发送周期
        7.报文长度
        8.信号名
        9.信号描述
        10.信号起始字节
        11.信号比特位
        12.可疑参数编号
        13.信号发送类型
        14.信号长度
        15.数据类型
        16.数据精度
        17.偏移量
        18.物理最小值
        19.物理最大值
        20.总线最小值
        21.总线最大值
        22.初始值
        23.无效值
        24.错误值
        25.非使能值
        26.单位
        27.信号值描述
        28.信号排列格式
        29.信号接收者
        30.报文发送节点
        31.读取属性
     */
    /*该软件转化的格式*/
     QVector<int> ReadMessageOrder={0,2,4,5,6,30,29};
     QVector<int> ReadSingalOrder={7,26,30,10,13,27,12,14,15,16,17,18,25,28};
     while(clomn<excel->vecDatas.size())
    {
        if(vecDatas[clomn][0]!="")//改行为报文行
        {
            Message message;
            // qDebug()<<vecDatas[clomn];
            for(int i=0;i<ReadMessageOrder.length();i++)
            {
                //读取报文行的信息
                SetInfo(ReadMessageOrder[i],vecDatas[clomn][i],&message,NULL);

            }
            clomn++;
            //接着读信号行的
            while(clomn<excel->vecDatas.size()&&vecDatas[clomn][0]==""&&clomn<excel->vecDatas.size())//循环条件依据对应的项目修改
            {
                Singal singal;
                for(int i=0;i<ReadSingalOrder.length();i++)
                {
                    //读取报文行的信息
                    SetInfo(ReadSingalOrder[i],vecDatas[clomn][i+ReadMessageOrder.length()],NULL,&singal);

                }
                message.singals.append(singal);
                clomn++;
            }
            messages.append(message);
        }
    }
      qDebug()<<"anlay ok";
 }


 /*
将读取不同的信息分为不同的步骤，方便修改顺序

*/
void Exceltodbc_sence::SetInfo(int step,QString info,Message *message,Singal *singal)
{
    switch (step) {
        case 0://报文名称
            if(message!=NULL)
            {
                message->MessageName=info;
            }
            break;
        case 1://报文类型
            if(message!=NULL)
            {

            }
            break;
        case 2://报文ID
            if(message!=NULL)
            {
                message->MessageId=info;
            }
            break;
        case 3://参数组编号
            if(message!=NULL)
            {

            }
            break;
        case 4://报文发送类型
            if(message!=NULL)
            {
                if(info=="Cycle")
                {
                    InsideFeature feature;
                    feature.Name="GenMsgCycleTime";
                    feature.Object="Message";
                    feature.SettingValue=info;
                    message->features.append(feature);
                }
            }
            break;
        case 5://报文发送周期
            if(message!=NULL)
            {
                if(message->features.length()!=0)
                {
                    for(int i=0;i<message->features.length();i++)
                    {
                        if(message->features[i].Name=="GenMsgCycleTime")
                        {
                            message->features[i].SettingValue=info;
                        }
                    }
                }
            }
            break;
        case 6://报文长度
            if(message!=NULL)
            {
                message->MessageSize=info;
            }
            break;
        case 7://信号名
            if(singal!=NULL)
            {
                singal->SignalName=info;
            }
            break;
        case 8://信号描述
            if(singal!=NULL)
            {
                singal->Description=info;
            }
            break;
        case 9://信号起始字节
            if(singal!=NULL)
            {
                //singal->StartBit=info;
            }
            break;
        case 10://信号比特位
            if(singal!=NULL)
            {
                singal->StartBit=info;
            }
            break;
        case 11://可疑参数编号
            if(message!=NULL)
            {

            }
            break;
        case 12://信号发送类型
            if(singal!=NULL)
            {
                //singal->=info;
            }
            break;
        case 13://信号长度
            if(singal!=NULL)
            {
                singal->SignalSize=info;
            }
            break;
        case 14://数据类型
            if(singal!=NULL)
            {
                singal->ValueType=info;
            }
            break;
        case 15://数据精度
            if(singal!=NULL)
            {
                singal->Factor=info;
            }
            break;
        case 16://偏移量
            if(singal!=NULL)
            {
                singal->Offset=info;
            }
            break;
        case 17://物理最小值
            if(singal!=NULL)
            {
                singal->Min=info;
            }
            break;
        case 18://物理最大值
            if(singal!=NULL)
            {
                singal->Max=info;
            }else if(singal!=NULL)
            {

            }
            break;
        case 19://总线最小值
            if(singal!=NULL)
            {
                //singal->Factor=info;
            }
            break;
        case 20://总线最大值
            if(singal!=NULL)
            {
                //singal->Factor=info;
            }
            break;
        case 21://初始值
            if(singal!=NULL)
            {
                //singal->Factor=info;
            }
            break;
        case 22://无效值
            if(singal!=NULL)
            {
                //singal->Factor=info;
            }
            break;
        case 23://错误值
            if(singal!=NULL)
            {
                //singal->Factor=info;
            }
            break;
        case 24://非使能值
            if(singal!=NULL)
            {
                //singal->Factor=info;
            }
            break;
        case 25://单位
            if(singal!=NULL)
            {
                singal->Unit=info;
            }
            break;
        case 26://信号值描述
            if(singal!=NULL)
            {
                singal->Description=info;
            }
            break;
        case 27://报文顺序
            if(singal!=NULL)
            {
                singal->ByteOrder=info;
            }
             break;
        case 28://报文发送者
            if(singal!=NULL)
            {
                singal->Receiver=info;
                if(Nodelist.indexOf(info)==-1)
                {
                    Nodelist.append(info);
                }
            }
            break;
        case 29://报文发送者
            if(message!=NULL)
            {
                message->Transmitter=info;
            }
            break;
        case 30://读取属性
            QList<QString> list=info.simplified().split(" ");
            Feature fea;
            InsideFeature infea;

            if(info!=""&&info!="\n"&&info!=" ")//不重复
            {
                qDebug()<<list;
                if(list[1]=="ENUM"){
                    fea.Name=list[0];fea.ValueType=list[1];fea.Min=list[2];
                    fea.DefaultValue=list[3];infea.SettingValue=list[4];
                }else if(list[2]=="STRING "){
                    fea.Name=list[0];fea.ValueType=list[1];

                }else{
                    fea.Name=list[0];fea.ValueType=list[1];fea.Min=list[2];fea.Max=list[3];
                    fea.DefaultValue=list[4];infea.SettingValue=list[5];
                }

                CopyStruct(&fea,&infea);
                qDebug()<<findfeaturebyname(list[0]);
                if(message!=NULL)//报文
                {

                    if(findfeaturebyname(list[0])==-1)
                    {
                       Messagefeatures.append(fea);
                    }
                    message->features.append(infea);
                }else if(singal!=NULL)//信号
                {
                    if(findfeaturebyname(list[0])==-1)
                    {
                       Singalfeatures.append(fea);
                    }
                    singal->features.append(infea);
                }
            }
            break;
    }
}

void Exceltodbc_sence::SaveDbc()
{
    //创建文件
    QFile file(QApplication::applicationDirPath() + "/ExcelToDbc.dbc");
   // qDebug()<<QApplication::applicationDirPath() + "/ExcelToDbc.dbc";
    if(file.open(QIODevice::ReadWrite|QIODevice::Text)){
       // qDebug()<<"create new file successfully";
    }else{
        //qDebug()<<"failed to create a new file!";
    }

    QTextStream out(&file);
    //VERSION
    QString version="VERSION \"\"\n\n\n";
    //NS
    QString NS="NS_:\n\tNS_DESC_\n\tCM_\n\tBA_DEF_\n\tBA_\n\tVAL_\n\tCAT_DEF_\n\tCAT_\n\tFILTER\n\tBA_DEF_DEF_\n\tEV_DATA_\n\tENVVAR_DATA_\n\t\
SGTYPE_\n\tSGTYPE_VAL_\n\tBA_DEF_SGTYPE_\n\tBA_SGTYPE_\n\tSIG_TYPE_REF_\n\tVAL_TABLE_\n\tSIG_GROUP_\n\tSIG_VALTYPE_\n\tSIGTYPE_VALTYPE_\n\tBO_TX_BU_\n\tBA_DEF_REL_\n\t\
BA_REL_\n\tBA_DEF_DEF_REL_\n\tBU_SG_REL_\n\tBU_EV_REL_\n\tBU_BO_REL_\n\tSG_MUL_VAL_\n\n";
    //BS
    QString BS="BS_:\n\n";
    //BU
    QString BU="BU_:";
    for(int i=0;i<Nodelist.length();i++)
    {
        if(Nodelist[i]!="Vector__XXX")
        {
            BU+=" ";
            BU+=Nodelist[i];
        }

    }
    BU+="\n";
    out<<version+NS+BS+BU;

    //BO
    //SG
    QList<QString> BO_SG;
    for(int i=0;i<messages.length();i++)
    {
        QString BO="BO_ ";

        BO+=(messages[i].MessageId+" ");
        BO+=(messages[i].MessageName+": ");
        BO+=(messages[i].MessageSize+" ");
        if(messages[i].Transmitter==""||messages[i].Transmitter=="Vector__XXX")
        {
            BO+=("Vector__XXX\n");
        }else{
            BO+=(messages[i].Transmitter+"\n");
        }
        //qDebug()<<"ok";
        for(int j=0;j<messages[i].singals.length();j++)
        {
            QString SG=" SG_ ";
            SG+=(messages[i].singals[j].SignalName+" : ");
            SG+=(messages[i].singals[j].StartBit+"|");
            SG+=(messages[i].singals[j].SignalSize+"@");
            if(messages[i].singals[j].ByteOrder=="Inter")
            {
                SG+=("1");
            }else
            {
                 SG+=("0");
            }
            if(messages[i].singals[j].ValueType=="Unsigned")
            {
                SG+=("+ (");
            }else
            {
                 SG+=("- (");
            }
            SG+=(messages[i].singals[j].Factor+","+messages[i].singals[j].Offset+") [");
            SG+=(messages[i].singals[j].Min+"|"+messages[i].singals[j].Max+"] ");
            SG+=("\""+messages[i].singals[j].Unit+"\" ");
            if(messages[i].singals[j].Receiver==""||messages[i].singals[j].Receiver=="Vector__XXX")
            {
                SG+=("Vector__XXX\n");
            }else{
                SG+=(messages[i].singals[j].Receiver+"\n");
            }

            BO+=SG;
        }
        BO_SG.append(BO+"\n");
    }


    for(int i=0;i<BO_SG.length();i++)
    {
        out<<BO_SG[i];
    }
    qDebug()<<"BO_SG ok";

    //BA_DEF_
    QList<QString> BA_DEF;
    for(int j=0;j<Singalfeatures.length();j++)
    {
        QString temp="";
        temp+=("BA_DEF_ SG_ ");
        if(Singalfeatures[j].Max==""||Singalfeatures[j].Max==" "||Singalfeatures[j].Max=="\n")
        {
            temp+=("\" " +Singalfeatures[j].Name+"\" "+Singalfeatures[j].ValueType+" "+Singalfeatures[j].Min+";");
        }else{
            temp+=("\" " +Singalfeatures[j].Name+"\" "+Singalfeatures[j].ValueType+" "+Singalfeatures[j].Min+" "+Singalfeatures[j].Max+";");
        }

        //qDebug()<<"BA_DEF_"+temp;
        BA_DEF.append(temp);
    }
    for(int j=0;j<Messagefeatures.length();j++)
    {
        QString temp="";
        temp+=("BA_DEF_ BO_ ");
        if(Messagefeatures[j].Max==""||Messagefeatures[j].Max==" "||Messagefeatures[j].Max=="\n")
        {
             temp+=("\"" +Messagefeatures[j].Name+"\" "+Messagefeatures[j].ValueType+" "+Messagefeatures[j].Min+";");
        }else{
            temp+=("\"" +Messagefeatures[j].Name+"\" "+Messagefeatures[j].ValueType+" "+Messagefeatures[j].Min+" "+Messagefeatures[j].Max+";");
        }

        //qDebug()<<"BA_DEF_"+temp;
        BA_DEF.append(temp);
    }
    //qDebug()<<BA_DEF[i];
    for(int i=0;i<BA_DEF.length();i++)
    {

         out<<(BA_DEF[i]+"\n");
    }
    qDebug()<<"BA_DEF ok";

    //BA_DEF_DEF_
    QList<QString> BA_DEF_DEF_;
    for(int j=0;j<Singalfeatures.length();j++)
    {
        if(Singalfeatures[j].DefaultValue!=""){
            QString temp="";
            temp+=("BA_DEF_DEF_ ");
            temp+=("\"" +Singalfeatures[j].Name+"\" "+Singalfeatures[j].DefaultValue+";");
            BA_DEF_DEF_.append(temp);
            qDebug()<<temp;
        }
    }
    for(int j=0;j<Messagefeatures.length();j++)
    {
        if(Messagefeatures[j].DefaultValue!=""){
            QString temp="";
            temp+=("BA_DEF_DEF_ ");
            temp+=("\"" +Messagefeatures[j].Name+"\" "+Messagefeatures[j].DefaultValue+";");
            BA_DEF_DEF_.append(temp);qDebug()<<temp;
        }
    }
    for(int i=0;i<BA_DEF_DEF_.length();i++)
    {
         out<<(BA_DEF_DEF_[i]+"\n");
    }
    qDebug()<<"BA_DEF_DEF_ ok";

    //BA_
    QList<QString> BA;
    for(int i=0;i<messages.length();i++)
    {
        for(int j=0;j<messages[i].features.length();j++)
        {
            QString temp="";
            temp+=("BA_ \""+messages[i].features[j].Name+"\" BO_ "+messages[i].MessageId+" "+messages[i].features[j].SettingValue+";");
            BA.append(temp);
        }
    }

    for(int i=0;i<messages.length();i++)
    {
        for(int j=0;j<messages[i].singals.length();j++)
        {
            for(int k=0;k<messages[i].singals[j].features.length();k++)
            {
                QString temp="";
                temp+=("BA_ \""+messages[i].singals[j].features[k].Name+"\" SG_ "+messages[i].MessageId+" "+messages[i].singals[j].SignalName+" "+messages[i].singals[j].features[k].SettingValue+";");
                BA.append(temp);
            }
        }
    }
    for(int i=0;i<BA.length();i++)
    {
         out<<(BA[i]+"\n");
    }
     qDebug()<<"BA_ ok";
    //VAL
    QList<QString> VAL;
    for(int i=0;i<messages.length();i++)
    {
        for(int j=0;j<messages[i].singals.length();j++)
        {
            if(messages[i].singals[j].Description!="")
            {
                //qDebug()<<messages[i].singals[j].Description;
                QString temp="";
                temp+=("VAL_ "+messages[i].MessageId+" "+messages[i].singals[j].SignalName+" ");
                QList<QString> list=messages[i].singals[j].Description.split("\n");
                for(int k=0;k<list.length();k++)
                {
                    temp+=(QString::number(list.length()-k-1)+" \""+list[k].mid(3,list[k].length()-3)+"\" ");
                }
                temp+=";";
                //qDebug()<<temp;
                VAL.append(temp);
            }
        }
    }
    //qDebug()<<VAL[i];
    for(int i=0;i<VAL.length();i++)
    {
         out<<(VAL[i]+"\n");
        // qDebug()<<VAL[i];
    }
    qDebug()<<"VAL ok";
    //关闭文件
    file.close();
}
int Exceltodbc_sence::findfeaturebyname(QString name)
{
    for(int i=0;i<Singalfeatures.length();i++)
    {
        if(Singalfeatures[i].Name==name){
            return i;
        }
    }
    for(int i=0;i<Messagefeatures.length();i++)
    {
        if(Messagefeatures[i].Name==name){
            return i;
        }
    }
    return -1;
}
void Exceltodbc_sence::CopyStruct(Feature *feature,InsideFeature *insidefeature)
{
    insidefeature->DefaultValue=feature->DefaultValue;
    insidefeature->ValueType=feature->ValueType;
    insidefeature->Object=feature->Object;
    insidefeature->Name=feature->Name;
    insidefeature->Min=feature->Min;
    insidefeature->Max=feature->Max;
}
