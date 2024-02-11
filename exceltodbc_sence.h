#ifndef EXCELTODBC_SENCE_H
#define EXCELTODBC_SENCE_H

#include <QMainWindow>
#include <QWidget>
#include <QTextEdit>
#include "excel.h"
#include "message_and_singal.h"
class Exceltodbc_sence: public QMainWindow
{
    Q_OBJECT
public:
    Exceltodbc_sence(QWidget *parent = nullptr);

private:
    Excel *excel;
    QTextEdit *textEdit;
    QList<Message> messages;
    QList<Feature> Singalfeatures;
    QList<Feature> Messagefeatures;
    QList<QString> Nodelist;
    void Dbc_Analyse();
    void SaveDbc();
    int findfeaturebyname(QString name);
    void CopyStruct(Feature *feature,InsideFeature *insidefeature);
    void SetInfo(int step,QString info,Message *message=NULL,Singal *singal=NULL);

};

#endif // EXCELTODBC_SENCE_H

