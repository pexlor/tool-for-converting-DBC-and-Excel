#ifndef MESSAGE_AND_SINGAL_H
#define MESSAGE_AND_SINGAL_H
#include <QWidget>
struct Feature{
    QString Object;
    QString Name;
    QString ValueType;
    QString Min;
    QString Max;
    QString DefaultValue;
};
struct InsideFeature : Feature{
     QString SettingValue;
};
struct Singal{
    QString SignalName;
    QString StartBit;
    QString SignalSize;
    QString ByteOrder;
    QString ValueType;
    QString Factor;
    QString Offset;
    QString Min;
    QString Max;
    QString Unit;
    QString Receiver;
    QString Description;
    QList<InsideFeature> features;
};
struct Message{
    QString MessageId;
    QString MessageName;
    QString MessageSize;
    QString Transmitter;
    QList<InsideFeature> features;
    QList<Singal> singals;
};

class Message_And_Singal
{
public:
    Message_And_Singal();
};

#endif // MESSAGE_AND_SINGAL_H
