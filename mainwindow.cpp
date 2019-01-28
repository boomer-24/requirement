#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
//    this->setWindowTitle("Формирование требования");
    this->setWindowTitle("Афонасиус");
    this->dir_ = false;
    this->file_ = false;    
    this->Initialize(QCoreApplication::applicationDirPath().append("/ini.xml"));
    this->ui->pushButton_3->setEnabled(false);
    QObject::connect(this->ui->lineEdit, SIGNAL(textChanged(QString)), this, SLOT(slotLineEditTextChanged(QString)));
    QObject::connect(this->ui->lineEdit_2, SIGNAL(textChanged(QString)), this, SLOT(slotLineEditCsvTextChanged(QString)));
}

MainWindow::~MainWindow()
{
    delete ui;
    this->SaveXml("ini.xml");
}

void MainWindow::FillCvsContent(QString _csvPath)
{    
    if (!_csvPath.isEmpty())
    {
        QFile csvFile(_csvPath);
        if (!csvFile.open(QIODevice::ReadOnly))
            qDebug() << _csvPath << " not open";
        else
        {
            bool checkTitle(false);
            QTextStream stream(&csvFile);
            int columnName = 0, columnAmount = 0;
            while (!stream.atEnd())
            {
                QString line = stream.readLine();
                if (!line.isEmpty())
                {
                    QStringList slist = line.split(";");
                    if (!checkTitle)
                        for (int i = 0; i < slist.size(); i++)
                        {
                            if (slist.at(i).contains("имен"))
                            {columnName = i; continue;}
                            if (slist.at(i).contains("оличеств"))
                            {columnAmount = i; checkTitle = true;}
                        }
                    else
                    {
                        QString s(slist.at(columnName));
                        if (!s.isEmpty())
                            this->vCsvContent_.push_back(QPair<QString, QString>(slist.at(columnName), slist.at(columnAmount)));
                    }
                }
            }
        }
    }
}

void MainWindow::Initialize(const QString &_xmlPath)
{
    QDomDocument domDoc;
    QFile file(_xmlPath);
    if (file.open(QIODevice::ReadOnly))
    {
        if (domDoc.setContent(&file))
        {
            QDomElement domElement = domDoc.documentElement();
            QDomNode domNode = domElement.firstChild();
            while(!domNode.isNull())
            {
                if (domNode.isElement())
                {
                    QDomElement domElement = domNode.toElement();
                    if (!domElement.isNull())
                    {
                        if (domElement.tagName() == "karpunin")
                        {
                            this->ui->lineEdit_3->setText(domElement.text());
                        } else if (domElement.tagName() == "number")
                        {
                            this->ui->spinBox->setValue(domElement.text().toInt());
                        }
                    }
                    domNode = domNode.nextSibling();
                }
            }
        }
    }
}

void MainWindow::SaveXml(const QString &_xmlPath)
{
    QFile file(_xmlPath);
    if (file.open(QIODevice::WriteOnly))
    {
        QXmlStreamWriter xmlWriter(&file);
        xmlWriter.setAutoFormatting(true);
        xmlWriter.writeStartDocument();
        xmlWriter.writeStartElement("general");

        xmlWriter.writeStartElement("karpunin");
        xmlWriter.writeCharacters(this->ui->lineEdit_3->text());
        xmlWriter.writeEndElement();

        xmlWriter.writeStartElement("number");
        xmlWriter.writeCharacters(QString::number(this->ui->spinBox->value()));
        xmlWriter.writeEndElement();

        xmlWriter.writeEndElement();
        xmlWriter.writeEndDocument();
        file.close();
    }
}

void MainWindow::on_pushButton_2_clicked()
{
    QString csvPath = QFileDialog::getOpenFileName(this, "Select csv-file", "C:/", "*.csv");
    this->ui->lineEdit_2->setText(csvPath);
}

void MainWindow::on_pushButton_clicked()
{
    QString dirPath(QFileDialog::getExistingDirectory(this, "Укажи папку", "C:/"));
    this->ui->lineEdit->setText(dirPath);
}

void MainWindow::slotLineEditTextChanged(QString _text)
{
    QDir dir(_text);
    if (!dir.exists() || _text.isEmpty())
    {
        this->dir_ = false;
        this->ui->pushButton_3->setEnabled(false);
    } else
    {
        this->dir_ = true;
        this->dirPath_ = this->ui->lineEdit->text();
        if (this->file_)
            this->ui->pushButton_3->setEnabled(true);
    }
}

void MainWindow::slotLineEditCsvTextChanged(QString _text)
{
    QFile csvFile(_text);
    if (!csvFile.exists() || _text.isEmpty())
    {
        this->file_ = false;
        this->ui->pushButton_3->setEnabled(false);
    }
    else
    {
        this->file_ = true;
        this->csvFile_ = _text;
        if (this->dir_)
            this->ui->pushButton_3->setEnabled(true);
    }
}

void MainWindow::on_pushButton_3_clicked()
{
    this->vCsvContent_.clear();
    this->FillCvsContent(this->csvFile_);
    Require r;
    r.setKarpunin(this->ui->lineEdit_3->text());
    for(int i = 0; i < this->vCsvContent_.size(); i++)
    {
        if (!this->vCsvContent_.at(i).first.isEmpty())
        {
            r.setNumber(this->ui->spinBox->value());
            r.CreateList(this->vCsvContent_.at(i).first, this->vCsvContent_.at(i).second);
            this->ui->spinBox->setValue(this->ui->spinBox->value() + 1);
            if (i < this->vCsvContent_.size() - 1)
                r.NewList();
        }
    }
    QString docxName = this->csvFile_.split("/").last();
    docxName.remove(QRegExp("\\.csv$"));
    QString dirPath = this->dirPath_;
    r.SaveList(dirPath.append("/").append(docxName));
}
