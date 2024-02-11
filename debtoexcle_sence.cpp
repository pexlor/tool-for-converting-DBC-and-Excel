#include "debtoexcle_sence.h"
#include <QPainter>
#include <QPushButton>
#include <QFileDialog>
#include <QDebug>
#include "excel.h"
#include <QDir>
#include <QMessageBox>
#include <QTextEdit>

#define TitleColor 150,200,100
#define MessageColor 50,200,100
#define SingalColor 255,255,255
DebToExcle_Sence::DebToExcle_Sence(QWidget *parent) : QMainWindow(parent)
{
    /*场景布置*/
    this->setFixedSize(500,500);
    this->setWindowIcon(QIcon(":/res/Set.bmp"));
    this->setWindowTitle("Dbc To Excel");

    QPushButton *ChooseBtn = new QPushButton("选择DBC文件",this);
    ChooseBtn->setFixedSize(150,80);
    ChooseBtn->move(30,this->height()/5);

    QPushButton *BeginBtn = new QPushButton("输出为Excel",this);
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
        QString filter="*.dbc"; //文件过滤器
        QString aFileName=QFileDialog::getOpenFileName(this,dlgTitle,curPath,filter);
        if (!aFileName.isEmpty()){
            DBCpath=aFileName;
            ReadDbc(DBCpath);
        }
    });

    connect(BeginBtn,&QPushButton::clicked,[=](){
        SaveMessage();
        textEdit->append("输出完毕请查看！");
    });
}


void DebToExcle_Sence::paintEvent(QPaintEvent *)
{
    QPainter painter(this);
    QPixmap pix;
    pix.load(":/res/OIP.jpg");
    painter.drawPixmap(0,0,this->width(),this->height(),pix);
    //画背景上的图标
}

void DebToExcle_Sence::ReadDbc(QString dbcpath)
{
    /*读取DBC文件并写入excel*/
    QFile file(dbcpath);
    file.open(QIODevice::ReadOnly);
    QString line;
    messages.clear();
    while(!file.atEnd())
    {

        line=file.readLine();
        /*读取报文信息和相应的信号信息*/
        /**************************/
        if(line.startsWith("BO_"))
        {
            Message message;
            QStringList MessageInfolist=line.split(" ");
            if(MessageInfolist.length()<5)
            {
                //提示错误
                break;
            }
            message.MessageId=MessageInfolist[1];
            message.MessageName=MessageInfolist[2];
            if(message.MessageName.right(1)==":")
            {
                message.MessageName.remove(message.MessageName.length()-1,1);
            }
            message.MessageSize=MessageInfolist[3];
            message.Transmitter=MessageInfolist[4];

            QString singalInfo;
            do
            {
                singalInfo=file.readLine();
                if(singalInfo.startsWith(" SG_"))
                {
                    /*这个地方可以搞个正则表达式来搞*/
                    Singal singal;
                    QStringList singallist=singalInfo.simplified().split(" ");
                    singal.SignalName=singallist[1];

                    int index1 = singallist[3].indexOf("|");
                    int index2 = singallist[3].indexOf("@");
                    singal.StartBit=singallist[3].mid(0,index1);
                    singal.SignalSize=singallist[3].mid(index1+1,index2-index1-1);
                    singal.ByteOrder=singallist[3].mid(index2+1,1);
                    singal.ValueType=singallist[3].mid(index2+2,1);

                    index1 = singallist[4].indexOf(",");
                    index2 = singallist[4].indexOf(")");
                    singal.Factor=singallist[4].mid(1,index1-1);
                    singal.Offset=singallist[4].mid(index1+1,index2-index1-1);

                    index1 = singallist[5].indexOf("|");
                    index2 = singallist[5].indexOf("]");
                    singal.Min=singallist[5].mid(1,index1-1);
                    singal.Max=singallist[5].mid(index1+1,index2-index1-1);

                    index1 = singallist[6].indexOf("\"");
                    singal.Unit=singallist[6].mid(index1+1,singallist[6].length()-2);
                    singal.Receiver=singallist[7];
                    message.singals.append(singal);
                    SignalCount++;
                }
            }while(singalInfo.startsWith(" SG_"));
            messages.append(message);
            MessageCount++;
        }else if(line.startsWith("BA_DEF_ "))//保存特征属性
        {
            struct Feature feature;
            QStringList Featureinfolist=line.simplified().split(" ");

            if(Featureinfolist[1]=="SG_"||Featureinfolist[1]=="BO_"||Featureinfolist[1]=="BU_")
            {

                feature.Object=Featureinfolist[1];
                feature.Name=Featureinfolist[2].mid(1,Featureinfolist[2].length()-2);

                if(Featureinfolist[3]=="HEX"||Featureinfolist[3]=="INT"||Featureinfolist[3]=="FLOAT")
                {

                    feature.ValueType=Featureinfolist[3];
                    feature.Min=Featureinfolist[4];
                    feature.Max=Featureinfolist[5].mid(0,Featureinfolist[5].length()-1);

                }else if(Featureinfolist[3]=="STRING")
                {
                    feature.ValueType=Featureinfolist[3];
                }else if(Featureinfolist[3]=="ENUM")
                {
                    feature.ValueType=Featureinfolist[3];
                    feature.Min=Featureinfolist[4].mid(0,Featureinfolist[4].length()-1);
                    qDebug()<<feature.Min;
                    //feature.Max=Featureinfolist[4].mid(Featureinfolist[4].indexOf(",")+2,Featureinfolist[4].length()-2);
                }

            }else
            {
                feature.Object="";
                feature.Name=Featureinfolist[1].mid(1,Featureinfolist[1].length()-2);

                if(Featureinfolist[3]=="HEX"||Featureinfolist[3]=="INT"||Featureinfolist[3]=="FLOAT")
                {
                    feature.ValueType=Featureinfolist[3];
                    feature.Min=Featureinfolist[4];
                    feature.Max=Featureinfolist[5].mid(0,Featureinfolist[5].length()-2);
                }else if(Featureinfolist[3]=="STRING")
                {
                    feature.ValueType=Featureinfolist[3];
                }else if(Featureinfolist[3]=="ENUM")
                {
                    feature.ValueType=Featureinfolist[3];
                    feature.Min=Featureinfolist[4].mid(0,Featureinfolist[4].length()-1);
                    qDebug()<<feature.Min;
                }
            }
           // qDebug()<<Featureinfolist[3];
            features.append(feature);
            featureCount++;
        }else if(line.startsWith("BA_DEF_REL_ "))
        {

        }else if(line.startsWith("BA_DEF_DEF_ "))//特征属性默认值
        {
            QStringList Featureinfolist=line.simplified().split(" ");

            int index_of_feature=FindFectureByName(Featureinfolist[1].mid(1,Featureinfolist[1].length()-2));

            if(index_of_feature!=-1)
            {
                features[index_of_feature].DefaultValue=Featureinfolist[2].mid(0,Featureinfolist[2].length()-1);
            }
        }else if(line.startsWith("BA_ "))//对应信号或报文特征属性的设定值
        {
            QStringList Featureinfolist=line.simplified().split(" ");
            int index_of_feature=FindFectureByName(Featureinfolist[1].mid(1,Featureinfolist[1].length()-2));//找到特征属性
            InsideFeature feature;

            if(Featureinfolist[2]=="BO_")
            {
                int index_of_message=FindMessageByID(Featureinfolist[3]);//找到对应的报文
                if(index_of_message!=-1&&index_of_feature!=-1)
                {
                   CopyStruct(&features[index_of_feature],&feature);
                   feature.SettingValue=Featureinfolist[4].mid(0,Featureinfolist[4].length()-1);
                   messages[index_of_message].features.append(feature);
                }
            }
            if(Featureinfolist[2]=="SG_")
            {
                int index_of_message=FindMessageByID(Featureinfolist[3]);//找到对应的报文
                int index_of_singal=FindSingalByName(Featureinfolist[4],index_of_message);//找到对应的信号

                if(index_of_message!=-1&&index_of_feature!=-1&&index_of_singal!=-1){
                    CopyStruct(&features[index_of_feature],&(feature));
                    feature.SettingValue=Featureinfolist[5].mid(0,Featureinfolist[5].length()-1);
                    messages[index_of_message].singals[index_of_singal].features.append(feature);

                }
            }
        }else if(line.startsWith("VAL_ "))
        {
            //提取description
            QVector<QString> extractedValues;
            int startPos = line.indexOf("\"");
            int endPos;
            while (startPos != -1) {
                endPos = line.indexOf("\"", startPos + 1);
                if (endPos != -1) {
                    QString extracted = line.mid(startPos + 1, endPos - startPos - 1);
                    extractedValues.push_back(extracted);
                    startPos = line.indexOf("\"", endPos + 1);
                } else {
                    break;
                }
            }

            //提取ID和Name
            QStringList Valueinfolist=line.simplified().split(" ");
            int index_of_message=FindMessageByID(Valueinfolist[1]);
            int index_of_singal=FindSingalByName(Valueinfolist[2],index_of_message);

            //保存信息
            if(index_of_message!=-1&&index_of_singal!=-1){
                for(int i=0;i<extractedValues.length();i++)
                {
                    messages[index_of_message].singals[index_of_singal].Description+=("0x"+QString::number(extractedValues.length()-1-i)+extractedValues[i]+"\n");
                }
            }
        }
        /**************************/
    }
    textEdit->append("读取成功！");
    file.close();
}


void DebToExcle_Sence::SaveMessage()
{
    Excel excle=Excel();
    QList<QString> list={"Message Name\n报文名称","Message ID\n报文标识符","Message Send Type\n报文发送类型","Message Cycle Time\n报文周期时间","Message Size\n报文长度","报文属性",
                         "Transmitter\n发送节点","Singal Name\n信号名称","Signa Value Description\n信号值描述","信号属性","Syart bit\n起始位","Length\n信号长度","Byte Order\n排列格式","Signal SendType\n信号发送类型",
                         "Value Type\n数据类型","Factor\n精度","Offset\n偏移量","Minimum\n最小值","Maximum\n最大值","Unit\n单位","Receiver\n接收者"};

    QString path =QDir::currentPath();
    QString fileName = path+"/DbcToExcel.xlsx";
    qDebug()<<fileName;
    excle.newExcel(fileName);
    setExcelTitle(excle,list);//设置标题
    setExcelMessage(excle);//写入内容
    excle.saveExcel(fileName);//保存文件
    excle.freeExcel();
}

void DebToExcle_Sence::setExcelTitle(Excel excle,QList<QString> ExcleTitle)
{
    int width[]={500,200,200,200,160,400,150,400,270,110,110,120,170,170,100,100,110,110,100,150};
    for(int i=1;i<=ExcleTitle.length();i++)
    {
       excle.setCellValue(1,i,ExcleTitle[i-1],QColor(TitleColor),(float)width[i-1]/9.21375);
    }
}


/*写入EXCEl*/
void DebToExcle_Sence::setExcelMessage(Excel excle)
{

    int index_row=2;
    int index_column=1;
    for(int i=1;i<=messages.length();i++)//写报文
    {
        index_column=1;
        excle.setCellValue(index_row,index_column++,messages[i-1].MessageName,QColor(MessageColor));
        excle.setCellValue(index_row,index_column++,messages[i-1].MessageId,QColor(MessageColor));
        int index_of_feature=FindFeatureByName("GenMsgCycleTime",i-1);
        if(index_of_feature!=-1)
        {
            excle.setCellValue(index_row,index_column++,"Cycle",QColor(MessageColor));
            excle.setCellValue(index_row,index_column++,messages[i-1].features[index_of_feature].SettingValue,QColor(MessageColor));
        }else
        {
            excle.setCellValue(index_row,index_column++,"NULL",QColor(MessageColor));
            excle.setCellValue(index_row,index_column++,"NULL",QColor(MessageColor));
        }
        excle.setCellValue(index_row,index_column++,messages[i-1].MessageSize,QColor(MessageColor));//报文长度
        QString message_atr="";
        qDebug()<<"ok";
        for(int k=0;k<messages[i-1].features.length();k++)
        {

            message_atr+=(messages[i-1].features[k].Name+" "+messages[i-1].features[k].ValueType+" "+messages[i-1].features[k].Min+" "+messages[i-1].features[k].Max+" "+messages[i-1].features[k].DefaultValue+" "+messages[i-1].features[k].SettingValue+"\n");

        }
        qDebug()<<"ok";
        excle.setCellValue(index_row,index_column++,message_atr,QColor(MessageColor));//报文属性
        excle.setCellValue(index_row,index_column++,messages[i-1].Transmitter,QColor(MessageColor));//发送节点
        index_row++;
        qDebug()<<"ok";
        for(int k=1;k<=messages[i-1].singals.length();k++)//写信号
        {
            index_column=8;
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].SignalName );
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].Description );//信号值描述
            QString sigal_atr="";
            //qDebug()<<messages[i-1].singals[k].features.length();
            for(int j=0;j<messages[i-1].singals[k-1].features.length();j++)
            {
                sigal_atr+=(messages[i-1].singals[k-1].features[j].Name+" "+messages[i-1].singals[k-1].features[j].ValueType+" "+messages[i-1].singals[k-1].features[j].Min+" "+messages[i-1].singals[k-1].features[j].Max+" "+messages[i-1].singals[k-1].features[j].DefaultValue+" "+messages[i-1].singals[k-1].features[j].SettingValue+"\n");

            }
            excle.setCellValue(index_row,index_column++,sigal_atr);//信号属性
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].StartBit );
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].SignalSize );
            if(messages[i-1].singals[k-1].ByteOrder=="1")
            {
                excle.setCellValue(index_row,index_column++,"Inter" );
            }else if(messages[i-1].singals[k-1].ByteOrder=="0")
            {
                excle.setCellValue(index_row,index_column++,"Motorola" );
            }

            if(index_of_feature!=-1)
            {
                excle.setCellValue(index_row,index_column++,"Cycle" );
            }else
            {
                excle.setCellValue(index_row,index_column++,"NULL" );
            }
            if(messages[i-1].singals[k-1].ValueType=="+")
            {
                excle.setCellValue(index_row,index_column++,"Unsigned" );
            }else
            {
                excle.setCellValue(index_row,index_column++,"Signed" );
            }
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].Factor );
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].Offset );
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].Min );
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].Max );
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].Unit );
            excle.setCellValue(index_row,index_column++,messages[i-1].singals[k-1].Receiver );
            index_row++;
        }
    }
}

int DebToExcle_Sence::FindFectureByName(QString name)
{
    for(int i=0;i<features.length();i++)
    {
        if(features[i].Name==name)
        {
            return i;
        }
    }
    return -1;
}

int DebToExcle_Sence::FindMessageByID(QString ID)
{
    for(int i=0;i<messages.length();i++)
    {
        if(messages[i].MessageId==ID)
        {
            return i;
        }
    }
    return -1;
}

int DebToExcle_Sence::FindSingalByName(QString name,int MessageIndex)
{
    for(int i=0;i<messages[MessageIndex].singals.length();i++)
    {
        if(messages[MessageIndex].singals[i].SignalName==name)
        {
            return i;
        }
    }
    return -1;
}

int DebToExcle_Sence::FindFeatureByName(QString name,int MessageIndex)
{
    for(int i=0;i<messages[MessageIndex].features.length();i++)
    {
        if(messages[MessageIndex].features[i].Name==name)
        {
            return i;
        }
    }
    return -1;
}

void DebToExcle_Sence::CopyStruct(Feature *feature,InsideFeature *insidefeature)
{
    insidefeature->DefaultValue=feature->DefaultValue;
    insidefeature->ValueType=feature->ValueType;
    insidefeature->Object=feature->Object;
    insidefeature->Name=feature->Name;
    insidefeature->Min=feature->Min;
    insidefeature->Max=feature->Max;
}
