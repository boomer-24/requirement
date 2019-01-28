#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include "require.h"
#include <QMainWindow>
#include <QFileDialog>
#include <QFile>
#include <QTextStream>
#include <QLabel>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();
    void FillCvsContent(QString _csvPath);
    void Initialize(const QString &_xmlPath);
    void SaveXml(const QString &_xmlPath);

private slots:
    void on_pushButton_clicked();
    void on_pushButton_2_clicked();
    void on_pushButton_3_clicked();
    void slotLineEditTextChanged(QString);
    void slotLineEditCsvTextChanged(QString);

private:
    Ui::MainWindow *ui;
    QVector<QPair<QString, QString>> vCsvContent_;
    QString dirPath_, csvFile_;        
    bool dir_, file_;    
};

#endif // MAINWINDOW_H
