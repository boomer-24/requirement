#ifndef REQUIRE_H
#define REQUIRE_H

#include <QAxObject>
#include <QDebug>
#include <QtXml>
#include <QDate>
#include <QFile>

class Require
{
private:
    QAxObject* wordApplication_;
    QAxObject* wordDocument_;
    QAxObject* activeDocument_;
    QAxObject* selection_;
    QAxObject* range_;
    QAxObject* tables_;
    QAxObject* table_;
    QAxObject* font_;

    QString name_, amount_, karpunin_;
    int number_;
    QMap<int, QString> mapMonth_;

public:
    Require();
    Require(QString _name, QString _amount);
    ~Require();

    void CreateList(QString _name, QString _amount);
    void NewList();
    void SaveList();
    void SaveList(QString _dirPath);
    void setNumber(int _number);
    void setKarpunin(const QString &_karpunin);
};

#endif // REQUIRE_H
