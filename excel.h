#ifndef EXCEL_H
#define EXCEL_H

#include <QObject>
#include <QAxObject>
#include <QMainWindow>
#include <QWidget>
class Excel
{
public:
    Excel();
    QVector<QVector<QString>> vecDatas;
    void newExcel(const QString &fileName);// 新建一个excel
    void appendSheet(const QString &sheetName,int cnt);// 增加一个worksheet

    void setCellValue(int row,int column,const QString &value,QColor color,int width);// 向Excel单元格中写入数据
    void setCellValue(int row,int column,const QString &value,QColor color);// 向Excel单元格中写入数据
    void setCellValue(int row,int column,const QString &value);// 向Excel单元格中写入数据

    void  readExcel(QString path,QString strSheetName);

    void saveExcel(const QString &fileName);// 保存excel
    void freeExcel();// 释放excel

private:

    QAxObject *pApplication;
    QAxObject *pWorkBooks;
    QAxObject *pWorkBook;
    QAxObject *pSheets;
    QAxObject *pSheet;
};

#endif // EXCEL_H
