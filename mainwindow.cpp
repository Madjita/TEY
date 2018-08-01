#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QtDebug>
#include <QSqlQueryModel>
#include <QSqlRecord>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    fileNames_S_R.append(ui->lineEdit->text());
    fileNames.append(ui->lineEdit_2->text());

    fileNames_XP_XS_XW_X.append(ui->lineEdit_3->text());
    fileNames_C_Z.append(ui->lineEdit_4->text());

    fileNames_BQ.append(ui->lineEdit_6->text());

    fileNames_DA.append(ui->lineEdit_7->text());


    fileNames_findMSWord.append( ui->lineEdit_8->text());

     fileNames_U.append(ui->lineEdit_9->text());

     fileNames_L.append(ui->lineEdit_10->text());

     fileNames_DD.append(ui->lineEdit_11->text());

     fileNames_TV.append(ui->lineEdit_12->text());


    InformLoading = new QLabel();

    InformLoading->setText("");

    ui->statusBar->addWidget(InformLoading);

    prBar = new QProgressBar();

    prBar->setVisible(false);

    ui->statusBar->addWidget(prBar);

    qDebug() << fileNames;

    word = new MYWORD(fileNames[0],fileNames_S_R[0],fileNames_XP_XS_XW_X[0],fileNames_C_Z[0],fileNames_BQ[0],fileNames_DA[0],fileNames_U[0],fileNames_L[0],fileNames_DD[0],fileNames_TV[0]);

    word->SetTemp(ui->lineEdit_5->text().toInt());



    connect(this,&MainWindow::begin,word,&MYWORD::Work);

    connect(word,&MYWORD::ChangeWork,this,&MainWindow::ChangeBar);

    connect(this,&MainWindow::findOnMSWord,word,&MYWORD::WorkFind);

    connect(word,&MYWORD::Part,this,&MainWindow::GetPart,Qt::DirectConnection);


    ui->lineEdit_13->setText(qApp->applicationDirPath());

    on_pushButton_12_clicked();


    bd = new BData();

    word->setBD(bd);


}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_clicked()
{


   InformLoading->setText("Создание карт: ");


  // word = new MYWORD(fileNames[0],fileNames_S_R[0],fileNames_XP_XS_XW_X[0],fileNames_C_Z[0],fileNames_BQ[0],fileNames_DA_DD[0]);


   word->SetTemp(ui->lineEdit_5->text().toInt());

   emit begin();


}

void MainWindow::on_pushButton_S_R_clicked()
{
    fileNames_S_R.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_S_R = dialog.selectedFiles();
        ui->lineEdit->setText(fileNames_S_R[0]);
    }


}

void MainWindow::on_pushButton_D_clicked()
{
    fileNames.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames = dialog.selectedFiles();


        word->FileDir = fileNames[0];

        ui->lineEdit_2->setText(fileNames[0]);
    }
}

void MainWindow::on_pushButton_2_clicked()
{
    fileNames_XP_XS_XW_X.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_XP_XS_XW_X = dialog.selectedFiles();

    ui->lineEdit_3->setText(fileNames_XP_XS_XW_X[0]);
    }
}

void MainWindow::on_pushButton_3_clicked()
{
    fileNames_C_Z.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_C_Z = dialog.selectedFiles();

    ui->lineEdit_4->setText(fileNames_C_Z[0]);
    }
}

void MainWindow::on_pushButton_4_clicked()
{
    fileNames_BQ.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_BQ = dialog.selectedFiles();

    ui->lineEdit_6->setText(fileNames_BQ[0]);
    }
}

void MainWindow::on_pushButton_5_clicked()
{
    fileNames_DA.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_DA = dialog.selectedFiles();

    ui->lineEdit_7->setText(fileNames_DA[0]);
    }
}

void MainWindow::ChangeBar(int max)
{

    qDebug () << QString::number(max);

    qDebug () << prBar->maximum();

    if(max != prBar->maximum())
    {
        prBar->setMaximum(max);

        prBar->setValue(0);

    }

    prBar->setValue(prBar->value()+1);

}

void MainWindow::GetPart(QString part)
{
    InformLoading->setText(part);
}

void MainWindow::on_pushButton_7_clicked()
{
    fileNames_findMSWord.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_findMSWord = dialog.selectedFiles();

    ui->lineEdit_8->setText(fileNames_findMSWord[0]);
    }
}

void MainWindow::on_pushButton_6_clicked()
{
    word->SetDirFindMSWord(fileNames_findMSWord[0]);

    emit findOnMSWord();
}

void MainWindow::on_pushButton_8_clicked()
{

    fileNames_U.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_U = dialog.selectedFiles();

        ui->lineEdit_9->setText(fileNames_U[0]);
    }


}

void MainWindow::on_pushButton_9_clicked()
{
    fileNames_L.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_L = dialog.selectedFiles();

        ui->lineEdit_10->setText(fileNames_L[0]);
    }
}

void MainWindow::on_pushButton_10_clicked()
{
    fileNames_DD.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_DD = dialog.selectedFiles();

        ui->lineEdit_11->setText(fileNames_DD[0]);
    }
}

void MainWindow::on_pushButton_11_clicked()
{
    fileNames_TV.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_TV = dialog.selectedFiles();

        ui->lineEdit_12->setText(fileNames_TV[0]);
    }
}

void MainWindow::on_pushButton_12_clicked()
{

   QStringList list = ui->lineEdit_13->text().split('/');

   if(list.count() <= 0)
   {
      list = ui->lineEdit_13->text().split('\\');
   }

   QString str ="";

   for(int i=0; i < list.count();i++)
   {
       if(i != list.count()-1)
       {
            list[i].append("/");
            list[i].append("/");
       }

       str += list[i];
   }

    qDebug () << str;


    word->saveDir = str;
}
