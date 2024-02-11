#include "excel.h"
#include <QFileInfo>
#include <QDir>
#include <QDebug>
Excel::Excel()
{

}

// 新建一个excel
void Excel::newExcel(const QString &fileName)
{
    pApplication = new QAxObject("Excel.Application");
    if (pApplication == nullptr) {
        qWarning("pApplication\n");
        return;
    }
    pApplication->dynamicCall("SetVisible(bool)",false);// false不显示窗体
    pApplication->setProperty("DisplayAlerts",false);// 不显示任何警告信息
    pWorkBooks = pApplication->querySubObject("Workbooks");
    QFile file(fileName);
    if (file.exists()) {
        pWorkBook = pWorkBooks->querySubObject("Open(const QString&)",fileName);
    } else {
        pWorkBooks->dynamicCall("Add");
        pWorkBook = pApplication->querySubObject("ActiveWorkBook");
    }
    pSheets = pWorkBook->querySubObject("Sheets");
    pSheet = pSheets->querySubObject("Item(int)",1);
}

// 增加一个worksheet
void Excel::appendSheet(const QString &sheetName,int cnt)
{
    QAxObject *pLastSheet = pSheets->querySubObject("Item(int)",cnt);
    pSheets->querySubObject("Add(QVariant)",pLastSheet->asVariant());
    pSheet = pSheets->querySubObject("Item(int)",cnt);
    pLastSheet->dynamicCall("Move(QVariant)",pSheet->asVariant());
    pSheet->setProperty("Name",sheetName);
}

// 向Excel单元格中写入数据
void Excel::setCellValue(int row,int column,const QString &value,QColor color)
{
    QAxObject *pRange = pSheet->querySubObject("Cells(int,int)",row,column);
    //QAxObject * cells = pRange->querySubObject("Columns");
    //->dynamicCall("AutoFit");
    pRange->setProperty("HorizontalAlignment",-4108);
    pRange->setProperty("VerticalAlignment",-4108);
    if(pRange!=NULL)
    {
        pRange->dynamicCall("Value",value);
    }
    QAxObject *interior=pRange->querySubObject("Interior");
    interior->setProperty("Color",color);
    QAxObject *boarder=pRange->querySubObject("Borders");
    boarder->setProperty("Color",QColor(0,0,0));
}

// 向Excel单元格中写入数据
void Excel::setCellValue(int row,int column,const QString &value,QColor color,int width)
{
    QAxObject *pRange = pSheet->querySubObject("Cells(int,int)",row,column);
    //QAxObject * cells = pRange->querySubObject("Columns");
    //->dynamicCall("AutoFit");
    pRange->setProperty("HorizontalAlignment",-4108);
    pRange->setProperty("VerticalAlignment",-4108);
    pRange->setProperty("ColumnWidth", width);  //设置单元格列宽
    if(pRange!=NULL)
    {
        pRange->dynamicCall("Value",value);
    }
    QAxObject *interior=pRange->querySubObject("Interior");
    interior->setProperty("Color",color);
    QAxObject *boarder=pRange->querySubObject("Borders");
    boarder->setProperty("Color",QColor(0,0,0));
}

// 向Excel单元格中写入数据
void Excel::setCellValue(int row,int column,const QString &value)
{
    QAxObject *pRange = pSheet->querySubObject("Cells(int,int)",row,column);
    if(pRange!=NULL)
    {
        pRange->dynamicCall("Value",value);
    }
    QAxObject *boarder=pRange->querySubObject("Borders");
    boarder->setProperty("Color",QColor(0,0,0));
}

void  Excel::readExcel(QString strPath,QString strSheetName)
{
    QFile file(strPath);
    if(!file.exists()){
        //qWarning() << "CExcelTool::loadExcel 路径错误，或文件不存在,路径为"<<strPath;
        return;
    }

    QAxObject *excel = new QAxObject("Excel.Application");//excel应用程序
    excel->dynamicCall("SetVisible(bool)", false); //true 表示操作文件时可见，false表示为不可见
    QAxObject *workbooks = excel->querySubObject("WorkBooks");//所有excel文件
    QAxObject *workbook = workbooks->querySubObject("Open(QString&)", strPath);//按照路径获取文件
    QAxObject * worksheets = workbook->querySubObject("WorkSheets");//获取文件的所有sheet页
    QAxObject * worksheet = worksheets->querySubObject("Item(QString)", strSheetName);//获取文件sheet页
    if(nullptr == worksheet){
        //qWarning()<<strSheetName<<"Sheet页不存在。";
        //return vecDatas;
    }
    QAxObject * usedrange = worksheet->querySubObject("UsedRange");//有数据的矩形区域

        QAxObject * rows = usedrange->querySubObject("Rows");//获取行数
        int nRows = rows->property("Count").toInt();
       // qDebug()<<"CExcelTool::loadExcel 文件数据行数为"<<nRows-1<<"，不含表头";
        if(nRows <= 1){
           // qWarning()<<"CExcelTool::loadExcel 无数据，跳过该文件";
            return;
        }
        QAxObject * columns = usedrange->querySubObject("Columns");//获取列数
        int nColumns = columns->property("Count").toInt();
       // qDebug()<<"CExcelTool::loadExcel 文件数据列数为"<<nColumns;

       // qInfo()<<"读可用数据...";
        for(int i = 2;i <= nRows;i++){//第一行默认为表头，从第二行读起
            QVector<QString> vecDataRow;
            for(int j = 1;j <= nColumns;j++){
                QAxObject *cell = worksheet->querySubObject("Cells(int,int)",i,j);
                QString strValue = cell->property("Value2").toString().trimmed();
                vecDataRow.push_back(strValue);
            }
            // qDebug()<<vecDataRow;
            vecDatas.push_back(vecDataRow);
        }
        //qInfo()<<"数据载入vec完毕...";
        //关闭文件
        workbook->dynamicCall("Close()");
        excel->dynamicCall("Quit()");
        if (excel)
        {
            delete excel;
            excel = NULL;
        }
}

// 保存excel
void Excel::saveExcel(const QString &fileName)
{
    pWorkBook->dynamicCall("SaveAs(const QString &)",QDir::toNativeSeparators(fileName));
}



// 释放excel
void Excel::freeExcel()
{
    if (pApplication != nullptr) {
        pApplication->dynamicCall("Quit()");
        delete pApplication;
        pApplication = nullptr;
    }
}


