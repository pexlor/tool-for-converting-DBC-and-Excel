#ifndef DEBTOEXCLE_SENCE_H
#define DEBTOEXCLE_SENCE_H

#include <QMainWindow>
#include <QWidget>
#include "excel.h"
#include <QTextEdit>
#include "message_and_singal.h"

class DebToExcle_Sence: public QMainWindow
{
    Q_OBJECT
public:
    int MessageCount=0;
    int SignalCount=0;
    int featureCount=0;
    QTextEdit *textEdit;
    explicit DebToExcle_Sence(QWidget *parent = nullptr);
    void paintEvent(QPaintEvent *);
private:
    QString DBCpath;
    QList<Message> messages;
    QList<Feature> features;
    int FindFectureByName(QString name);
    int FindMessageByID(QString ID);
    int FindSingalByName(QString name,int MessageIndex);
    int FindFeatureByName(QString name,int MessageIndex);
    void CopyStruct(Feature *feature,InsideFeature *insidefeature);
    void ReadDbc(QString DBCpath);
    void setExcelTitle(Excel excle,QList<QString> ExcleTitle);
    void setExcelMessage(Excel excle);
    void SaveMessage();
};

#endif // DEBTOEXCLE_SENCE_H
