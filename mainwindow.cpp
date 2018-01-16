#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QtDebug>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    fileNames_S_R.append(ui->lineEdit->text());
    fileNames.append(ui->lineEdit_2->text());

    fileNames_XP_XS_XW.append(ui->lineEdit_3->text());
    fileNames_C_Z.append(ui->lineEdit_4->text());

    fileNames_BQ.append(ui->lineEdit_6->text());

    fileNames_DA_DD.append(ui->lineEdit_7->text());


    fileNames_findMSWord.append( ui->lineEdit_8->text());




    InformLoading = new QLabel();

    InformLoading->setText("");

    ui->statusBar->addWidget(InformLoading);

    prBar = new QProgressBar();

    prBar->setVisible(false);

    ui->statusBar->addWidget(prBar);


    word = new MYWORD(fileNames[0],fileNames_S_R[0],fileNames_XP_XS_XW[0],fileNames_C_Z[0],fileNames_BQ[0],fileNames_DA_DD[0]);


    word->SetTemp(ui->lineEdit_5->text().toInt());

    word->moveToThread(new QThread());


   // connect(word->thread(),&QThread::started,word,&MYWORD::Work);

    connect(this,&MainWindow::begin,word,&MYWORD::Work);

    connect(word,&MYWORD::ChangeWork,this,&MainWindow::ChangeBar);

    connect(this,&MainWindow::findOnMSWord,word,&MYWORD::WorkFind);

    connect(word,&MYWORD::Part,this,&MainWindow::GetPart,Qt::DirectConnection);

    word->thread()->start();

}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pushButton_clicked()
{


   InformLoading->setText("Создание карт: ");


  // word = new MYWORD(fileNames[0],fileNames_S_R[0],fileNames_XP_XS_XW[0],fileNames_C_Z[0],fileNames_BQ[0],fileNames_DA_DD[0]);


  // word->SetTemp(ui->lineEdit_5->text().toInt());

  // word->moveToThread(new QThread());


  // connect(word->thread(),&QThread::started,word,&MYWORD::Work);


  // word->Work();

  // word->thread()->start();

   emit begin();

   // ui->lineEdit->setText(fileNames[0]);
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

    ui->lineEdit_2->setText(fileNames[0]);
    }
}

void MainWindow::on_pushButton_2_clicked()
{
    fileNames_XP_XS_XW.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_XP_XS_XW = dialog.selectedFiles();

    ui->lineEdit_3->setText(fileNames_XP_XS_XW[0]);
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
    fileNames_DA_DD.clear();

    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::AnyFile);

    if (dialog.exec())
    {
        fileNames_DA_DD = dialog.selectedFiles();

    ui->lineEdit_7->setText(fileNames_DA_DD[0]);
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
