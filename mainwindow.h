#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QFileDialog>
#include <QThread>

#include <word.h>

#include <QProgressBar>
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

     QStringList fileNames,fileNames_S_R,fileNames_XP_XS_XW,fileNames_C_Z,fileNames_BQ,fileNames_DA_DD,fileNames_findMSWord;

     MYWORD* word ;

     QLabel* InformLoading;
     QProgressBar* prBar;


signals:
     void begin();

     void findOnMSWord();

private slots:
    void on_pushButton_clicked();

    void on_pushButton_S_R_clicked();

    void on_pushButton_D_clicked();

    void on_pushButton_2_clicked();

    void on_pushButton_3_clicked();



    void on_pushButton_4_clicked();

    void on_pushButton_5_clicked();

    void ChangeBar(int);

    void GetPart(QString);

    void on_pushButton_7_clicked();

    void on_pushButton_6_clicked();

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
