#include "word.h"

#include <QtDebug>



MYWORD::MYWORD(QString _FileDir, QString _FileDir_S_R,QString _FileDir_XP_XS_XW,QString _FileDir_C_Z,QString _FileDir_BQ,QString _FileDir_DA_DD,QObject *parent) : QObject(parent),
    FileDir(_FileDir),
    FileDir_S_R(_FileDir_S_R),
    FileDir_XP_XS_XW(_FileDir_XP_XS_XW),
    FileDir_C_Z(_FileDir_C_Z),
    FileDir_BQ(_FileDir_BQ),
    FileDir_DA_DD(_FileDir_DA_DD)
{

}

void MYWORD::SetTemp(int R)
{
    temp =R;
}

void MYWORD::SetDirFindMSWord(QString data)
{
    FileDir_FindMSWord = data;
}


void MYWORD::Work()
{
    // OpenWord();

    OpenWord_Perechen();
}


void MYWORD::WorkFind()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",FileDir_FindMSWord); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");




    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");



    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    k = ActiveDocument_2->querySubObject("Tables")->property("Count").toInt();


    qDebug () << "K = " << k;



    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=1;i <= k;i++)
    {

        Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(i));


        // flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 2); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 2); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



        QString text =  CellRange_2->property("Text").toString();


        Find_E.append(text);


        text = CellRange_2_3->property("Text").toString();

        Find_EName.append(text);

        Find_Data_1.append(QStringList());
        Find_Data_2.append(QStringList());

        //Cбор данных

        auto columns = Tables_2->querySubObject("Columns")->property("Count").toInt();

        auto rows = Tables_2->querySubObject("Rows")->property("Count").toInt();


        if(FileDir_FindMSWord.split('/').last() == "XPXSXW.docx")
        {


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 4); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 5); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 5); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 2); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 6); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 6); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////





            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 3); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



            text =  CellRange_2->property("Text").toString();


            Find_E.append(text);


            text = CellRange_2_3->property("Text").toString();

            Find_EName.append(text);

            Find_Data_1.append(QStringList());
            Find_Data_2.append(QStringList());

            //Cбор данных

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 6); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 7); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 7); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 8); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 8); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 4); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



            text =  CellRange_2->property("Text").toString();


            Find_E.append(text);


            text = CellRange_2_3->property("Text").toString();

            Find_EName.append(text);

            Find_Data_1.append(QStringList());
            Find_Data_2.append(QStringList());

            //Cбор данных

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 8); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 9); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 9); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 10); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 10); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 10); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 10); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 11); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 11); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 17, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 23, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 24, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 25, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 26, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 27, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 28, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////

            // Производим поиск собранных данных


            qDebug () << Find_E;
            qDebug () << Find_EName;

            qDebug () << Find_Data_1;

            qDebug () << Find_Data_2;

        }


        if(FileDir_FindMSWord.split('/').last() == "R.docx")
        {


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 4); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 4); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 4); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 3); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 5); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 5); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 4); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////





            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 3); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



            text =  CellRange_2->property("Text").toString();


            Find_E.append(text);


            text = CellRange_2_3->property("Text").toString();

            Find_EName.append(text);

            Find_Data_1.append(QStringList());
            Find_Data_2.append(QStringList());

            //Cбор данных

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 6); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 6); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 6); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 5); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 7); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 7); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 6); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 4); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");



            text =  CellRange_2->property("Text").toString();


            Find_E.append(text);


            text = CellRange_2_3->property("Text").toString();

            Find_EName.append(text);

            Find_Data_1.append(QStringList());
            Find_Data_2.append(QStringList());

            //Cбор данных

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 8); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 8); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 8); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 7); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_1[Find_E.count()-1].append(text);


            //////////////////////////////////////////////////////////////////////////////////////


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9); // (ячейка 4:4)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);


            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 5, 9); // (ячейка 5:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 6, 9); // (ячейка 6:5)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 7, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 8, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 9, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 10, 10); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 11, 10); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 12, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 13, 8); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 14, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 15, 9); // (ячейка 7:6)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            text = CellRange_2->property("Text").toString();

            Find_Data_2[Find_E.count()-1].append(text);

            /////////////////////////////////////////////////////

            // Производим поиск собранных данных


            qDebug () << Find_E;
            qDebug () << Find_EName;

            qDebug () << Find_Data_1;

            qDebug () << Find_Data_2;

        }



        qDebug ()  << "========================================================";

    }


    QStringList Send_Find_E_bekap,result;

    Send_Find_E_bekap = Find_EName;


    bool flagApp=true;

    bool flagApp2=true;


    int f = 0;
    int save_sov_i = 0;



//    for(int i=0;i < Find_EName.count();i++)
//    {
//        if(result.count() < 1)
//        {
//            flagApp = true;

//            result.append(Find_EName[i]);
//            Send_Find_E.append(QStringList());
//            Send_Find_E[result.count()-1].append(Find_E[i]);
//            Send_Find_Data_1.append(Find_Data_1[i]);
//            Send_Find_Data_2.append(Find_Data_2[i]);

//        }
//        else
//        {

//            int res = result.count();

//            for(int j=0; j < res;j++)
//            {
//                flagApp = false;

//                if(Find_EName[i] == result[j])
//                {
//                    flagApp = true;

//                    break;
//                }

//            }

//            if(flagApp == false)
//            {
//                result.append(Find_EName[i]);
//                Send_Find_E.append(QStringList());
//                Send_Find_E[result.count()-1].append(Find_E[i]);
//                Send_Find_Data_1.append(Find_Data_1[i]);
//                Send_Find_Data_2.append(Find_Data_2[i]);
//            }
//            else
//            {
//                Send_Find_E[result.count()-1].append(Find_E[i]);
//                Send_Find_Data_1[result.count()-1].append(Find_Data_1[i]);
//                Send_Find_Data_2[result.count()-1].append(Find_Data_2[i]);

//            }

//        }


//    }

    QString first;
    do
    {

        flagApp = true;
        first = Find_EName[0];
        result.append(Find_EName[0]);
        Send_Find_E.append(QStringList());
        Send_Find_E[result.count()-1].append(Find_E[0]);
        Send_Find_Data_1.append(Find_Data_1[0]);
        Send_Find_Data_2.append(Find_Data_2[0]);

        for(int i=1;i < Find_EName.count();i++)
        {
            if(Find_EName[i] == first)
            {
                flagApp = false;
                Find_EName.removeAt(i);
                Send_Find_E[result.count()-1].append(Find_E[i]);
                Find_E.removeAt(i);
                Send_Find_Data_1[result.count()-1].append(Find_Data_1[i]);
                Find_Data_1.removeAt(i);
                Send_Find_Data_2[result.count()-1].append(Find_Data_2[i]);
                Find_Data_2.removeAt(i);
                i--;
            }
        }

        Find_EName.removeAt(0);
        Find_E.removeAt(0);
        Find_Data_1.removeAt(0);
        Find_Data_2.removeAt(0);

        first = "";




    }while(Find_EName.count() > 1);

    qDebug ()  << "========================================================";

    qDebug () << result;

    qDebug ()  << "======================111111===========================";

    qDebug () << Send_Find_Data_1;

    qDebug ()  << "====================2222222222============================";

    qDebug () << Send_Find_Data_2;

    qDebug ()  << "========================================================";

    qDebug () << Send_Find_E;

    qDebug ()  << "========================================================";



    bool flag_find_sovpad = false;

    QStringList result_2,Send_Find_Data_1_eshe,Send_Find_Data_2_eshe;
    QList<QStringList> Send_Find_E_eshe;

    for(int i=0 ;i < Send_Find_E.count();i++)
    {
        if(Send_Find_E[i].count() > 0)
        {

            for(int j=0; j < Send_Find_Data_1[i].count();j++)
            {

                lol2.append(Send_Find_Data_1[i].value(j));

                if(((j%11) == 0 ) && ((j > 0)&& (j <= 11)))
                {

                    list.append(lol2);
                    lol2.clear();
                }
                else
                {
                    qDebug() << QString::number(j%12);

                    if(((j%12) == 11 ) && (j > 11))
                    {
                        list.append(lol2);
                        lol2.clear();
                    }
                }

            }

            for(int j=0; j < Send_Find_Data_2[i].count();j++)
            {

                lol2_2.append(Send_Find_Data_2[i].value(j));

                if(((j%11) == 0 ) && ((j > 0)&& (j <= 11)))
                {

                    list2.append(lol2_2);
                    lol2_2.clear();
                }
                else
                {
                    qDebug() << QString::number(j%12);

                    if(((j%12) == 11 ) && (j > 11))
                    {
                        list2.append(lol2_2);
                        lol2_2.clear();
                    }
                }

            }

            //подумать !!!!
            for(int j=0;j < list.count()-1;j++)
            {
                if(  (list[j] != list[j+1]) ||  ( list2[j] != list2[j+1]))
                {
                  //  Send_Find_Data_1.replace(i,list[j]);
                  //  Send_Find_Data_2.replace(i,list2[j]);

//                    auto list_copy = list;

//                    QStringList first = list_copy[0];
//                    Send_Find_E.append(QStringList());

//                    do
//                    {
//                        for(int k =0;k < list_copy.count();k++)
//                        {
//                            if(list_copy[k] == first)
//                            {


//                                result.append(result[i]);

//                                Send_Find_E.last().append(Send_Find_E[i].value(j+1));
//                                Send_Find_E[i].removeAt(j+1);
//                                Send_Find_Data_1[result.count()-1].append(Find_Data_1[i]);
//                                Find_Data_1.removeAt(i);
//                                Send_Find_Data_2[result.count()-1].append(Find_Data_2[i]);
//                                Find_Data_2.removeAt(i);
//                                list_copy.removeAt(k);
//                                k--;
//                            }
//                        }

//                    }while(list_copy.count() < 1);

                  //  Send_Find_E.append(QStringList());
                  //  Send_Find_E.last().append(Send_Find_E[i].value(j+1));
                  //  Send_Find_E[i].removeAt(j+1);

                    Send_Find_E_eshe.append(QStringList());
                    Send_Find_E_eshe.last().append(Send_Find_E[i].value(j+1));
                    Send_Find_E[i].removeAt(j+1);

                    result_2.append(result[i]);

                  //  list.removeAt(j+1);
                 //   Send_Find_Data_1.append(list[j+1]);
                 //   Send_Find_Data_2.append(list2[j+1]);

                    list.removeAt(j+1);
                    Send_Find_Data_1_eshe.append(list[j+1]);
                    Send_Find_Data_2_eshe.append(list2[j+1]);

                    flag_find_sovpad = true;


                    break;
                }
                else
                {
//                    Send_Find_Data_1.replace(i,list[j]);
//                    Send_Find_Data_2.replace(i,list2[j]);
                }
            }


            list.clear();
            list2.clear();

            flag_find_sovpad = false;

        }
    }


    qDebug ()  << "========================================================";

    qDebug () << result;

    qDebug ()  << "======================111111===========================";

    qDebug () << Send_Find_Data_1;

    qDebug ()  << "====================2222222222============================";

    qDebug () << Send_Find_Data_2;

    qDebug ()  << "========================================================";

    qDebug () << Send_Find_E;

    qDebug ()  << "========================================================";


    ActiveDocument_2->dynamicCall("Close (boolean)", false);

    WordApplication_2->dynamicCall("Quit (void)");


    /////////////////////////////////////END////////////////////////////////////////////////////////////////////////////


    if(FileDir_FindMSWord.split('/').last() == "XPXSXW.docx")
    {

        WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

        // WordApplication_2->setProperty("Visible", 1);

        WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

        WordDocuments_2->querySubObject( "Open(%T)",FileDir_XP_XS_XW); //D:\\11111\\One.docx


        ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



        // ActiveDocument_2->querySubObject( "Range()" )->querySubObject("Copy()");


        ActiveDocument_2->querySubObject("Tables(1)")->querySubObject( "Range()" )->querySubObject("Copy()");






        selection_2 = WordApplication_2->querySubObject("Selection()");


        qDebug() <<"Send_Find_E.count()/3 = " << QString::number(Send_Find_E.count()%3);



        if((((Send_Find_E.count()-1)%3) > 0)  && (((Send_Find_E.count()-1)%3) !=0))
        {
            for(int i=1; i < (Send_Find_E.count()/3)+1;i++)
            {

                selection_2->dynamicCall("EndKey(wdStory)");

                selection_2->dynamicCall("InsertBreak()");


                selection_2->querySubObject( "Paste()");

            }
        }
        else
        {
            for(int i=1; i < (Send_Find_E.count()/3);i++)
            {

                selection_2->dynamicCall("EndKey(wdStory)");

                selection_2->dynamicCall("InsertBreak()");

                selection_2->querySubObject( "Paste()");

            }
        }



        /////////////////////////////////////////////////////


        flag =0;


        k = 1;




        selection_2->dynamicCall("HomeKey(wdStory)");


        qDebug () << "K = " << k;

        Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


        selection_2->dynamicCall("HomeKey(wdStory)");


        QString text;

        for(int i =0 ; i < Send_Find_E.count();i++)
        {

            flag++;

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            if(Send_Find_E[i].count() < 1)
            {
                CellRange_2->dynamicCall("InsertAfter(Text)", Send_Find_E[i].value(0));
            }
            else
            {
                QString str = "";

                for(int j=0;j < Send_Find_E[i].count();j++)
                {
                    if(j != Send_Find_E[i].count()-1)
                    {
                        str +=Send_Find_E[i].value(j).split(0x000d).first()+", ";
                    }
                    else
                    {
                        str +=Send_Find_E[i].value(j);
                    }
                }
                CellRange_2->dynamicCall("InsertAfter(Text)", str);
            }

            CellRange_2_3->dynamicCall("InsertAfter(Text)", result[i]);

            switch (flag) {

            case 1:
            {

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 4); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 5); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 5); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 2); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }



                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 6); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 6); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();
                if(text != Send_Find_Data_2[i].value(2))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////


                //////////////////////////////////////////////////////////////////////////////////////

                break;
            }
            case 2:
            {

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 6); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 7); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 7); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }


                //////////////////////////////////////////////////////////////////////////////////////



                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 8); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 8); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 9); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////

                break;
            }
            case 3:
            {


                   StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 8); // (ячейка 4:4)
                   CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(0))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                    }


                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 9); // (ячейка 5:5)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(1))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 9); // (ячейка 6:5)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(2))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 10); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(3))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 10); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(4))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(5))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(6))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(7))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 7); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(8))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 7); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(9))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                    }


                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 7); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(10))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 6); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = StartCell_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(11))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                    }


                //    //////////////////////////////////////////////////////////////////////////////////////


                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9); // (ячейка 4:4)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(0))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                    }


                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 10); // (ячейка 5:5)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(1))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 10); // (ячейка 6:5)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(2))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 11); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(3))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 11); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(4))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 17, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(5))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(6))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 24, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(7))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(8))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 26, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(9))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 27, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(10))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 28, 7); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(11))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                    }

                //    /////////////////////////////////////////////////////
                break;
            }

            }


            if(flag == 3)
            {
                flag =0;

                k++;
                if(k > (Send_Find_E.count()/3))
                {

                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                    qDebug () << "K = " << k;
                }

            }
        }





        //Сохранить pdf
        //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//XPXSXW" ,"17");//fileName.split('.').first()

        ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//XPXSXWRelase");


        ActiveDocument_2->dynamicCall("Close (boolean)", false);

        WordApplication_2->dynamicCall("Quit (void)");

    }


    if(FileDir_FindMSWord.split('/').last() == "R.docx")
    {

        WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

        // WordApplication_2->setProperty("Visible", 1);

        WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

        WordDocuments_2->querySubObject( "Open(%T)",FileDir_S_R); //D:\\11111\\One.docx


        ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



        // ActiveDocument_2->querySubObject( "Range()" )->querySubObject("Copy()");


        ActiveDocument_2->querySubObject("Tables(1)")->querySubObject( "Range()" )->querySubObject("Copy()");






        selection_2 = WordApplication_2->querySubObject("Selection()");


        qDebug() <<"Send_Find_E.count()/3 = " << QString::number(Send_Find_E.count()%3);



        if((((Send_Find_E.count()-1)%3) > 0)  && (((Send_Find_E.count()-1)%3) !=0))
        {
            for(int i=1; i < (Send_Find_E.count()/3)+1;i++)
            {

                selection_2->dynamicCall("EndKey(wdStory)");

                selection_2->dynamicCall("InsertBreak()");


                selection_2->querySubObject( "Paste()");

            }
        }
        else
        {
            for(int i=1; i < (Send_Find_E.count()/3);i++)
            {

                selection_2->dynamicCall("EndKey(wdStory)");

                selection_2->dynamicCall("InsertBreak()");

                selection_2->querySubObject( "Paste()");

            }
        }



        /////////////////////////////////////////////////////


        flag =0;


        k = 1;




        selection_2->dynamicCall("HomeKey(wdStory)");


        qDebug () << "K = " << k;

        Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


        selection_2->dynamicCall("HomeKey(wdStory)");


        QString text;

        for(int i =0 ; i < Send_Find_E.count();i++)
        {

            flag++;

            StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
            CellRange_2 = StartCell_2->querySubObject("Range()");

            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            if(Send_Find_E[i].count() < 1)
            {
                CellRange_2->dynamicCall("InsertAfter(Text)", Send_Find_E[i].value(0));
            }
            else
            {
                QString str = "";

                for(int j=0;j < Send_Find_E[i].count();j++)
                {
                    if(j != Send_Find_E[i].count()-1)
                    {
                        str +=Send_Find_E[i].value(j).split(0x000d).first()+", ";
                    }
                    else
                    {
                        str +=Send_Find_E[i].value(j);
                    }
                }
                CellRange_2->dynamicCall("InsertAfter(Text)", str);
            }

            CellRange_2_3->dynamicCall("InsertAfter(Text)", result[i]);

            switch (flag) {

            case 1:
            {

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 4); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 4); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 4); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 5); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }



                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 5); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 5); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();
                if(text != Send_Find_Data_2[i].value(2))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 3); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 4); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////


                //////////////////////////////////////////////////////////////////////////////////////

                break;
            }
            case 2:
            {

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 6); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 6); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 6); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(2))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(3))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(4))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(5))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(6))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(7))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(8))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 5); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(9))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(10))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_1[i].value(11))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                }


                //////////////////////////////////////////////////////////////////////////////////////



                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 7); // (ячейка 4:4)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(0))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 7); // (ячейка 5:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(1))
                {

                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 7); // (ячейка 6:5)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(2))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(3))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(4))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(5))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(6))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 8); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(7))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(8))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                }


                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 6); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(9))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(10))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                }

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 7); // (ячейка 7:6)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                text = CellRange_2_3->property("Text").toString();

                if(text != Send_Find_Data_2[i].value(11))
                {
                    CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                }

                //    /////////////////////////////////////////////////////

                break;
            }
            case 3:
            {


                   StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 8); // (ячейка 4:4)
                   CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(0))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(0));
                    }


                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 8); // (ячейка 5:5)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(1))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(1));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 8); // (ячейка 6:5)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(2))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(2));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(3))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(3));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(4))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(4));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(5))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(5));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(6))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(6));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(7))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(7));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(8))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(8));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 7); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(9))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(9));
                    }


                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(10))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(10));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 8); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = StartCell_2_3->property("Text").toString();

                    if(text != Send_Find_Data_1[i].value(11))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_1[i].value(11));
                    }


                //    //////////////////////////////////////////////////////////////////////////////////////


                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 4, 9); // (ячейка 4:4)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(0))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(0));
                    }


                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 5, 9); // (ячейка 5:5)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(1))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(1));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 6, 9); // (ячейка 6:5)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(2))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(2));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 7, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(3))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(3));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 8, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(4))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(4));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 9, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(5))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(5));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 10, 10); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(6))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(6));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 11, 10); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(7))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(7));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 12, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(8))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(8));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 13, 7); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(9))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(9));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(10))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(10));
                    }

                    StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 9); // (ячейка 7:6)
                    CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                    text = CellRange_2_3->property("Text").toString();

                    if(text != Send_Find_Data_2[i].value(11))
                    {
                        CellRange_2_3->dynamicCall("InsertAfter(Text)", Send_Find_Data_2[i].value(11));
                    }

                //    /////////////////////////////////////////////////////
                break;
            }

            }


            if(flag == 3)
            {
                flag =0;

                k++;
                if(k > (Send_Find_E.count()/3))
                {

                    qDebug () << "Конец ; K = " << k;
                    break;
                }
                else
                {
                    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                    qDebug () << "K = " << k;
                }

            }
        }





        //Сохранить pdf
        //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//XPXSXW" ,"17");//fileName.split('.').first()

        ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//RRelase");


        ActiveDocument_2->dynamicCall("Close (boolean)", false);

        WordApplication_2->dynamicCall("Quit (void)");

    }



    result.clear();

    Send_Find_Data_1.clear();

    Send_Find_Data_2.clear();

    Find_E.clear();
    Find_EName.clear();
    Find_Data_1.clear();
    Find_Data_2.clear();

    // То что нужно записать
    Send_Find_E.clear();
    Send_Find_EName.clear();
    Send_Find_Data_1.clear();
    Send_Find_Data_2.clear();

    Send_Find_Data_1_1.clear();
    Send_Find_Data_2_2.clear();

    list.clear();
    lol2.clear();





}



void MYWORD::OpenWord()
{
    QAxObject* WordApplication = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication->setProperty("Visible", 1);

    QAxObject* WordDocuments = WordApplication->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments->querySubObject( "Open(%T)",FileDir); //D:\\11111\\One.docx


    QAxObject* ActiveDocument = WordApplication->querySubObject("ActiveDocument()");



    QAxObject *selection2 = WordApplication->querySubObject("Selection()");


    QAxObject* Tables = selection2->querySubObject("Tables(1)");



    QAxObject* StartCell  = Tables->querySubObject("Cell(Row, Column)", 6, 2); // (ячейка C1)
    QAxObject* CellRange = StartCell->querySubObject("Range()");



    //CellRange->dynamicCall("InsertAfter(Text)", "НУ");


    //    StartCell = Tables->querySubObject("Cell(Row, Column)", 8, 3);

    //    CellRange = StartCell->querySubObject("Range()");



    //    auto lol =  CellRange->property("Text");

    //    qDebug () << lol.toString();

    auto columns = Tables->querySubObject("Columns")->property("Count").toInt();

    auto rows = Tables->querySubObject("Rows")->property("Count").toInt();

    qDebug () << "Колонки = " << columns;

    qDebug () <<"Строки = " << rows;


    //////////////////////////////////////////////////////////////////////////////
    int count_find = 0;

    for(int i=1; i <  rows;i++)
    {
        for(int j=1; j < columns; j++)
        {

            StartCell = Tables->querySubObject("Cell(Row, Column)", i, j);

            CellRange = StartCell->querySubObject("Range()");

            QString text =  CellRange->property("Text").toString();

            if((text[0] == "R") && (j == 2))
            {
                count_find++;

            }
        }
    }

    qDebug () << QString::number(count_find);

    ///////////////////////////////////////////////////////////////////////////////


    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",FileDir_S_R); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject( "Range()" )->querySubObject("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject( "Range()" )->querySubObject("Copy()");







    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug() <<"count_find/3 = " << QString::number(count_find/3);



    for(int i=0; i < (count_find/3);i++)
    {

        selection_2->dynamicCall("EndKey(wdStory)");

        selection_2->dynamicCall("InsertBreak()");


        selection_2->querySubObject( "Paste()");

    }



    //    QAxObject* Tables_2 = ActiveDocument_2->querySubObject("Tables(1)");



    //    QAxObject* StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 2); // (ячейка C1)
    //    QAxObject* CellRange_2 = StartCell_2->querySubObject("Range()");

    //    QAxObject* StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 2); // (ячейка C1)
    //    QAxObject* CellRange_2_3 = StartCell_2_3->querySubObject("Range()");




    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=1; i <  rows;i++)
    {
        for(int j=1; j < columns; j++)
        {

            StartCell = Tables->querySubObject("Cell(Row, Column)", i, j);

            CellRange = StartCell->querySubObject("Range()");

            QString text =  CellRange->property("Text").toString();

            if((text[0] == "R") && (j == 2))
            {

                flag++;

                StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
                CellRange_2 = StartCell_2->querySubObject("Range()");

                StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
                CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

                CellRange_2->dynamicCall("InsertAfter(Text)", text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();

                CellRange_2_3->dynamicCall("InsertAfter(Text)", text);



                if(flag == 3)
                {
                    flag =0;

                    k++;
                    if(k > (count_find/3)+1 )
                    {

                        qDebug () << "Конец ; K = " << k;
                        break;
                    }
                    else
                    {
                        Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                        qDebug () << "K = " << k;
                    }

                }

            }
        }
    }






    //Сохранить pdf
    ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//Good" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//Good");
    ActiveDocument_2->dynamicCall("Close (boolean)", false);
    ActiveDocument->dynamicCall("Close (boolean)", false);

    WordApplication->dynamicCall("Quit (void)");
    WordApplication_2->dynamicCall("Quit (void)");
}

void MYWORD::OpenWord_Perechen()
{

    Part("Открытие документа : " + FileDir);

    WordApplication = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord




    //  WordApplication->setProperty("Visible", 1); //Показать (Открыть) окно MSWord с документом

    WordDocuments = WordApplication->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments->querySubObject( "Open(%T)",FileDir); //D:\\11111\\One.docx


    ActiveDocument = WordApplication->querySubObject("ActiveDocument()"); // Сделать документ активным



    selection2 = WordApplication->querySubObject("Selection()");  // Создать класс Области страницы


    Tables = selection2->querySubObject("Tables(1)"); // Выбираем 1 таблицу в документе


    StartCell  = Tables->querySubObject("Cell(Row, Column)", 6, 2); // (ячейка C1)

    CellRange = StartCell->querySubObject("Range()"); // Область выбранной ячейки

    columns = Tables->querySubObject("Columns")->property("Count").toInt();

    rows = Tables->querySubObject("Rows")->property("Count").toInt();

    qDebug () << "Колонки = " << columns;

    qDebug () <<"Строки = " << rows;


    Part("Открыт : " + FileDir + " Количество Колонок: " + QString::number(columns) + " Строк: " +  QString::number(rows));


    Findelements_Perechen();

}

////////////////////////////////////////////////////////////////

void MYWORD::Findelements_Perechen()
{

    R.clear();      //отчистка резисторы
    RName.clear();  //отчистка имя резисторов


    C_Z.clear();    //отчистка конденсаторы и фильтры
    C_ZName.clear();  //отчистка имя конденсаторов

    XP_XS_XW.clear();  //отчистка Вилка
    XP_XS_XWName.clear(); //отчистка ИмяВилки

    DA_DD.clear();
    DA_DDName.clear();

    BQ.clear();
    BQName.clear();



    int count_find = 0;

    QString text;

    emit ChangeWork(rows*columns);


    Part("Ищим Элименты... Колонок: " + QString::number(columns) + " Строк: " +  QString::number(rows));

    for(int i=1; i <  rows;i++)
    {

        for(int j=1; j < columns; j++)
        {


            emit ChangeWork(rows*columns);

            StartCell = Tables->querySubObject("Cell(Row, Column)", i, j);

            CellRange = StartCell->querySubObject("Range()");

            text =  CellRange->property("Text").toString();


            //Ищим R (резисторы)
            if((text[0] == "R") && (j == 2))
            {
                count_find++;

                R.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                RName.append(text);
                break;

            }

            //Ищим C (конденсаторы)
            if((text[0] == "C") && (j == 2)) //С
            {
                C_Z.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                C_ZName.append(text);
                break;
            }

            //Ищим Z (фильтры)
            if((text[0] == "Z") && (j == 2))
            {
                C_Z.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                C_ZName.append(text);
                break;
            }

            //Ищим XP (вилка)
            if(((text[0] == "X") && (text[1] == "P")) && (j == 2))
            {
                XP_XS_XW.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                XP_XS_XWName.append(text);
                break;
            }

            //Ищим XS (Розетка)
            if(((text[0] == "X") && (text[1] == "S")) && (j == 2))
            {
                XP_XS_XW.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                XP_XS_XWName.append(text);
                break;
            }

            //Ищим XW (Вилка)
            if(((text[0] == "X") && (text[1] == "W")) && (j == 2))
            {
                XP_XS_XW.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                XP_XS_XWName.append(text);
                break;
            }

            //Ищим BQ (Резонатор)
            if(((text[0] == "B") && (text[1] == "Q")) && (j == 2))
            {
                BQ.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                BQName.append(text);
                break;
            }

            //Ищим DA (Микросхема)
            if(((text[0] == "D") && (text[1] == "A")) && (j == 2))
            {
                DA_DD.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                DA_DDName.append(text);
                break;
            }

            //Ищим DD (Микросхема)
            if(((text[0] == "D") && (text[1] == "D")) && (j == 2))
            {
                DA_DD.append(text);

                StartCell = Tables->querySubObject("Cell(Row, Column)", i, j+1);

                CellRange = StartCell->querySubObject("Range()");

                text =  CellRange->property("Text").toString();


                DA_DDName.append(text);
                break;
            }

        }
    }

    Part("Поиск завершен. Закрытие документа.");

    qDebug () << QString::number(count_find);

    qDebug () << R;

    qDebug () << RName;


    qDebug () << "=============================";

    qDebug () << C_Z.count();

    qDebug () << C_Z;

    qDebug () << C_ZName;

    qDebug () << "=============================";

    qDebug () << XP_XS_XW.count();

    qDebug () << XP_XS_XW;

    qDebug () << XP_XS_XWName;

    qDebug () << "=============================";

    qDebug () << BQ.count();

    qDebug () << BQ;

    qDebug () << BQName;

    qDebug () << "=============================";

    qDebug () << DA_DD.count();

    qDebug () << DA_DD;

    qDebug () << DA_DDName;


    ActiveDocument->dynamicCall("Close (boolean)", false);

    WordApplication->dynamicCall("Quit (void)");


    CreatShablon();

}

void MYWORD::CreatShablon()
{

    Part("Создание шаблона с XP XS XW.");

    if(XP_XS_XW.count() > 0)
    {
        CreatShablon_XP_XS_XW();
    }

    Sleep(2000);

    Part("Создание шаблона с С Z.");

    if(C_Z.count() > 0)
    {
        CreatShablon_C_Z();
    }

    Sleep(2000);

    Part("Создание шаблона с R.");

    if(R.count() > 0)
    {
        CreatShablon_R();
    }

    Sleep(2000);

    Part("Создание шаблона с BQ.");

    if(BQ.count() > 0)
    {
        CreatShablon_BQ();
    }

    Sleep(2000);

    Part("Создание шаблона с DA DD.");

    if(DA_DD.count() > 0)
    {
        CreatShablon_DA_DD();
    }



    R.clear();      //отчистка резисторы
    RName.clear();  //отчистка имя резисторов


    C_Z.clear();    //отчистка конденсаторы и фильтры
    C_ZName.clear();  //отчистка имя конденсаторов

    XP_XS_XW.clear();  //отчистка Вилка
    XP_XS_XWName.clear(); //отчистка ИмяВилки

    DA_DD.clear();
    DA_DDName.clear();

    BQ.clear();
    BQName.clear();

    Part("Шаблоны созданны.");


}

void MYWORD::CreatShablon_R()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    // WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",FileDir_S_R); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject( "Range()" )->querySubObject("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject( "Range()" )->querySubObject("Copy()");







    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug() <<"R.count()/3 = " << QString::number(R.count()/3);



    //    for(int i=1; i < (R.count()/3);i++)
    //    {

    //        selection_2->dynamicCall("EndKey(wdStory)");

    //        selection_2->dynamicCall("InsertBreak()");


    //        selection_2->querySubObject( "Paste()");

    //    }


    if((R.count()%3) > 0 )
    {
        for(int i=1; i < (R.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }
    else
    {
        for(int i=1; i < (R.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }



    //    QAxObject* Tables_2 = ActiveDocument_2->querySubObject("Tables(1)");



    //    QAxObject* StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 2); // (ячейка C1)
    //    QAxObject* CellRange_2 = StartCell_2->querySubObject("Range()");

    //    QAxObject* StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 2); // (ячейка C1)
    //    QAxObject* CellRange_2_3 = StartCell_2_3->querySubObject("Range()");




    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=1;i < R.count();i++)
    {
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", R[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", RName[i]);

        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 6); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 8); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }




        if(flag == 3)
        {
            flag =0;

            k++;
            if(k > (R.count()/3)+1 )
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                qDebug () << "K = " << k;
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//R" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//R");


    ActiveDocument_2->dynamicCall("Close (boolean)", false);

    WordApplication_2->dynamicCall("Quit (void)");
}

////////////////////////////////////////////////////


void MYWORD::CreatShablon_C_Z()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    // WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",FileDir_C_Z); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject( "Range()" )->querySubObject("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject( "Range()" )->querySubObject("Copy()");







    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug() <<"R.count()/3 = " << QString::number(C_Z.count()/3);



    //    for(int i=1; i < (C_Z.count()/3);i++)
    //    {

    //        selection_2->dynamicCall("EndKey(wdStory)");

    //        selection_2->dynamicCall("InsertBreak()");


    //        selection_2->querySubObject( "Paste()");

    //    }

    if((C_Z.count()%3) > 0 )
    {
        for(int i=1; i < (C_Z.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }
    else
    {
        for(int i=1; i < (C_Z.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }


    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < C_Z.count();i++)
    {
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", C_Z[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", C_ZName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 6); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 15, 8); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }


        if(flag == 3)
        {
            flag =0;

            k++;
            if(k > (C_Z.count()/3)+1 )
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                qDebug () << "K = " << k;
            }

        }

    }





    //Сохранить pdf
    //  ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//CZ" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//CZ");


    ActiveDocument_2->dynamicCall("Close (boolean)", false);

    WordApplication_2->dynamicCall("Quit (void)");
}

void MYWORD::CreatShablon_XP_XS_XW()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    // WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",FileDir_XP_XS_XW); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");



    // ActiveDocument_2->querySubObject( "Range()" )->querySubObject("Copy()");


    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject( "Range()" )->querySubObject("Copy()");






    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");




    qDebug() <<"XP_XS_XW.count()/3 = " << QString::number(XP_XS_XW.count()/3);



    //    for(int i=1; i < (XP_XS_XW.count()/3);i++)
    //    {

    //        selection_2->dynamicCall("EndKey(wdStory)");

    //        selection_2->dynamicCall("InsertBreak()");


    //        selection_2->querySubObject( "Paste()");

    //    }

    if((XP_XS_XW.count()%3) > 0 )
    {
        for(int i=1; i < (XP_XS_XW.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }
    else
    {
        for(int i=1; i < (XP_XS_XW.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }



    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < XP_XS_XW.count();i++)
    {
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", XP_XS_XW[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", XP_XS_XWName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 5); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 25, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }



        if(flag == 3)
        {
            flag =0;

            k++;
            if(k > (XP_XS_XW.count()/3)+1 )
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                qDebug () << "K = " << k;
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//XPXSXW" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//XPXSXW");


    ActiveDocument_2->dynamicCall("Close (boolean)", false);

    WordApplication_2->dynamicCall("Quit (void)");
}

void MYWORD::CreatShablon_BQ()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    //  WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",FileDir_BQ); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject( "Range()" )->querySubObject("Copy()");


    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");


    qDebug() <<"XP_XS_XW.count()/3 = " << QString::number(BQ.count()/3);



    //    for(int i=1; i < (BQ.count()/3);i++)
    //    {

    //        selection_2->dynamicCall("EndKey(wdStory)");

    //        selection_2->dynamicCall("InsertBreak()");


    //        selection_2->querySubObject( "Paste()");

    //    }


    if((BQ.count()%3) > 0 )
    {
        for(int i=1; i < (BQ.count()/3)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }
    else
    {
        for(int i=1; i < (BQ.count()/3);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }



    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < BQ.count();i++)
    {
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", BQ[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", BQName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 3); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 5); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 3:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 14, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }

        }



        if(flag == 3)
        {
            flag =0;

            k++;
            if(k > (BQ.count()/3)+1 )
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                qDebug () << "K = " << k;
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//BQ" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//BQ");


    ActiveDocument_2->dynamicCall("Close (boolean)", false);

    WordApplication_2->dynamicCall("Quit (void)");
}

void MYWORD::CreatShablon_DA_DD()
{
    QAxObject* WordApplication_2 = new QAxObject("Word.Application"); // Создаю интерфейс к MSWord

    //  WordApplication_2->setProperty("Visible", 1);

    QAxObject* WordDocuments_2 = WordApplication_2->querySubObject( "Documents()" ); // Получаю интерфейсы к его подобъекту "коллекция открытых документов":

    WordDocuments_2->querySubObject( "Open(%T)",FileDir_DA_DD); //D:\\11111\\One.docx


    QAxObject* ActiveDocument_2 = WordApplication_2->querySubObject("ActiveDocument()");

    ActiveDocument_2->querySubObject("Tables(1)")->querySubObject( "Range()" )->querySubObject("Copy()");


    QAxObject *selection_2 = WordApplication_2->querySubObject("Selection()");


    qDebug() <<"DA_DD.count()/2 = " << QString::number(DA_DD.count()/2);

    qDebug() <<"DA_DD.count()/2%10 = " << QString::number(DA_DD.count()%2);



    if((DA_DD.count()%2) > 0 )
    {
        for(int i=1; i < (DA_DD.count()/2)+1;i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }
    else
    {
        for(int i=1; i < (DA_DD.count()/2);i++)
        {

            selection_2->dynamicCall("EndKey(wdStory)");

            selection_2->dynamicCall("InsertBreak()");


            selection_2->querySubObject( "Paste()");

        }
    }






    /////////////////////////////////////////////////////


    QAxObject* Tables_2,*StartCell_2,*CellRange_2,*StartCell_2_3,*CellRange_2_3;

    int flag =0;

    int k = 1;


    qDebug () << "K = " << k;

    Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));


    selection_2->dynamicCall("HomeKey(wdStory)");



    for(int i=0;i < DA_DD.count();i++)
    {
        flag++;

        StartCell_2  = Tables_2->querySubObject("Cell(Row, Column)", 1, 1+flag); // (ячейка C1)
        CellRange_2 = StartCell_2->querySubObject("Range()");

        StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 2, 1+flag); // (ячейка C1)
        CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

        CellRange_2->dynamicCall("InsertAfter(Text)", DA_DD[i]);

        CellRange_2_3->dynamicCall("InsertAfter(Text)", DA_DDName[i]);


        //Темпиратура


        switch (flag) {

        case 1:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 4); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }
        case 2:
        {
            StartCell_2_3  = Tables_2->querySubObject("Cell(Row, Column)", 23, 7); // (ячейка C1)
            CellRange_2_3 = StartCell_2_3->querySubObject("Range()");

            CellRange_2_3->dynamicCall("InsertAfter(Text)", QString::number(temp));
            break;
        }


        }



        if(flag == 2)
        {
            flag =0;

            k++;
            if(k > (DA_DD.count()/2)+1 )
            {

                qDebug () << "Конец ; K = " << k;
                break;
            }
            else
            {
                Tables_2 = ActiveDocument_2->querySubObject("Tables(%T)",QString::number(k));

                qDebug () << "K = " << k;
            }

        }

    }





    //Сохранить pdf
    //   ActiveDocument_2->dynamicCall("ExportAsFixedFormat (const QString&,const QString&)","D://11111//1//DADD" ,"17");//fileName.split('.').first()

    ActiveDocument_2->dynamicCall("SaveAs2 (const QString&)", "D://11111//1//DADD");


    ActiveDocument_2->dynamicCall("Close (boolean)", false);

    WordApplication_2->dynamicCall("Quit (void)");
}


