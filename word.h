#ifndef WORD_H
#define WORD_H

#include <QObject>


#include <QAxObject>
#include <QAxWidget>
#include <QAxBase>

#include <windows.h>


class MYWORD : public QObject
{
    Q_OBJECT
public:
    explicit MYWORD(QString FileDir, QString FileDir_S_R,QString FileDir_XP_XS_XW_X,QString FileDir_C_Z, QString FileDir_BQ,QString FileDir_DA,QString FileDir_U,QString FileDir_L,QString FileDir_DD,QString FileDir_TV,QObject *parent = 0);

    QString FileDir,FileDir_S_R,FileDir_XP_XS_XW_X,FileDir_C_Z,FileDir_BQ,FileDir_DA,FileDir_DD,FileDir_FindMSWord;

    QString FileDir_U, FileDir_L, FileDir_TV;


    QList<QAxObject*> WordApplicationShablonList; // Шаблоны

    QAxObject* WordApplication, // Создаю интерфейс к MSWord  перечню

                *WordDocuments,  // Класс документа перечня
                *ActiveDocument, // Сделать документ активным
                *selection2,     // Создать класс Области страницы
                *Tables,         // Выбираем 1 таблицу в документе
                *StartCell,     // ячейка
                *CellRange;     // Область выбранной ячейки



    int temp;


    int columns; //Колонки в Перечне

    int rows; // Количество строчек в Перечне


    QStringList R; //резисторы
    QStringList RName; //имя резисторов


    QStringList C_Z;  //конденсаторы и фильтры
    QStringList C_ZName;  //имя конденсаторов


    //КАРТА   РАБОЧИХ   РЕЖИМОВ   ЭЛЕКТРИЧЕСКИХ   СОЕДИНЕНИЙ,   ПРОВОДОВ   И   КАБЕЛЕЙ
    QStringList XP_XS_XW_X;  //Вилка
    QStringList XP_XS_XW_XName; //ИмяВилки

    //КАРТА   РАБОЧИХ   РЕЖИМОВ   КВАРЦЕВЫХ   РЕЗОНАТОРОВ,   КВАРЦЕВЫХ   МИКРОГЕНЕРАТОРОВ,   ПЬЕЗОЭЛЕКТРИЧЕСКИХ
    //И ЭЛЕКТРОМЕХАНИЧЕСКИХ   ФИЛЬТРОВ   И   ЛИНИЙ   ЗАДЕРЖКИ   НА   ПОВЕРХНОСТНЫХ   АКУСТИЧЕСКИХ   ВОЛНАХ
    QStringList BQ;
    QStringList BQName;

    QStringList DD;
    QStringList DDName;

    QStringList DA;
    QStringList DAName;


    //РЕЖИМОВ   ВТОРИЧНЫХ   ИСТОЧНИКОВ   ПИТАНИЯ (ФОРМА  82/83)
    QStringList U;
    QStringList UName;


    QStringList L;
    QStringList LName;

    QStringList TV;
    QStringList TVName;



    //Берем элимент и название с данными

    QStringList Find_E;
    QStringList Find_EName;
    QList<QStringList> Find_Data_1;
    QList<QStringList> Find_Data_2;

    // То что нужно записать
    QList<QStringList> Send_Find_E;
    QStringList Send_Find_EName;
    QList<QStringList> Send_Find_Data_1;
    QList<QStringList> Send_Find_Data_2;

    QList<QStringList>  Send_Find_Data_1_1;
    QList<QStringList>  Send_Find_Data_2_2;

    QList<QStringList> list;
    QList<QStringList> list2;
    QStringList lol2;
    QStringList lol2_2;




signals:
    void ChangeWork(int);

    void Part(QString);

public slots:


    void SetTemp(int);

    void SetDirFindMSWord(QString);

    void OpenWord();

    void OpenWord_Perechen();

    void Findelements_Perechen();


    void CreatShablon();

    void CreatShablon_R();

    void CreatShablon_C_Z();

    void CreatShablon_XP_XS_XW_X();

    void CreatShablon_BQ();

    void CreatShablon_DA();

    void CreatShablon_DD();

    void CreatShablon_U();

    void CreatShablon_L();

    void CreatShablon_TV();

    void Work();

    void WorkFind();

};

#endif // WORD_H
