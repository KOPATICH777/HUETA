#include <iostream>
#include <OpenXLSX.hpp>
#include <sstream>
int toInt(const std::string& str) {
    std::stringstream ss(str);
    int res=0;
    ss >> res;
    return res;
}
int main() {
    try {
        // Открываем исходный файл
        OpenXLSX::XLDocument sourceDoc;
        sourceDoc.open("D:\\Test\\up01_03_021.xlsx"); // Замените "SourceFile.xlsx" на путь к вашему исходному файлу
        auto sourceWks = sourceDoc.workbook().worksheet("Sheet1"); // Замените "Sheet1" на имя вашего листа
        
        // Читаем данные из исходного файла
        // Для 1 предмета
        std::string data1 = sourceWks.cell("D17").value().getString();
        std::string data2 = sourceWks.cell("E17").value().getString();

        std::string data3 = sourceWks.cell("K17").value().getString();
        std::string data4 = sourceWks.cell("O17").value().getString();
        std::string data5 = sourceWks.cell("S17").value().getString();

        std::string data6 = sourceWks.cell("AR17").value().getString();
        std::string data7 = sourceWks.cell("AV17").value().getString();
        std::string data8 = sourceWks.cell("AZ17").value().getString();

        std::string data9 = sourceWks.cell("BY17").value().getString();
        std::string data10 = sourceWks.cell("CC17").value().getString();
        std::string data11 = sourceWks.cell("CG17").value().getString();

        // Для 2 предмета
        std::string data12 = sourceWks.cell("D18").value().getString();
        std::string data13 = sourceWks.cell("E18").value().getString();

        std::string data14 = sourceWks.cell("K18").value().getString();
        std::string data15 = sourceWks.cell("O18").value().getString();
        std::string data16 = sourceWks.cell("S18").value().getString();

        std::string data17 = sourceWks.cell("AR18").value().getString();
        std::string data18 = sourceWks.cell("AV18").value().getString();
        std::string data19 = sourceWks.cell("AZ18").value().getString();

        std::string data20 = sourceWks.cell("BY18").value().getString();
        std::string data21 = sourceWks.cell("CC18").value().getString();
        std::string data22 = sourceWks.cell("CG18").value().getString();

        // Для 3 предмета
        std::string data23 = sourceWks.cell("D19").value().getString();
        std::string data24 = sourceWks.cell("E19").value().getString();

        std::string data25 = sourceWks.cell("K19").value().getString();
        std::string data26 = sourceWks.cell("O19").value().getString();
        std::string data27 = sourceWks.cell("S19").value().getString();

        std::string data28 = sourceWks.cell("AR19").value().getString();
        std::string data29 = sourceWks.cell("AV19").value().getString();
        std::string data30 = sourceWks.cell("AZ19").value().getString();

        std::string data31 = sourceWks.cell("BY19").value().getString();
        std::string data32 = sourceWks.cell("CC19").value().getString();
        std::string data33 = sourceWks.cell("CG19").value().getString();

        // Для 4 предмета
        std::string data34 = sourceWks.cell("D20").value().getString();
        std::string data35 = sourceWks.cell("E20").value().getString();

        std::string data36 = sourceWks.cell("K20").value().getString();
        std::string data37 = sourceWks.cell("O20").value().getString();
        std::string data38 = sourceWks.cell("S20").value().getString();

        std::string data39 = sourceWks.cell("AR20").value().getString();
        std::string data40 = sourceWks.cell("AV20").value().getString();
        std::string data41 = sourceWks.cell("AZ20").value().getString();

        std::string data42 = sourceWks.cell("BY20").value().getString();
        std::string data43 = sourceWks.cell("CC20").value().getString();
        std::string data44 = sourceWks.cell("CG20").value().getString();

        // Для 5 предмета
        std::string data45 = sourceWks.cell("D21").value().getString();
        std::string data46 = sourceWks.cell("E21").value().getString();

        std::string data47 = sourceWks.cell("K21").value().getString();
        std::string data48 = sourceWks.cell("O21").value().getString();
        std::string data49 = sourceWks.cell("S21").value().getString();

        std::string data50 = sourceWks.cell("AR21").value().getString();
        std::string data51 = sourceWks.cell("AV21").value().getString();
        std::string data52 = sourceWks.cell("AZ21").value().getString();

        std::string data53 = sourceWks.cell("BY21").value().getString();
        std::string data54 = sourceWks.cell("CC21").value().getString();
        std::string data55 = sourceWks.cell("CG21").value().getString();

        // Для 6 предмета
        std::string data56 = sourceWks.cell("D22").value().getString();
        std::string data57 = sourceWks.cell("E22").value().getString();

        std::string data58 = sourceWks.cell("K22").value().getString();
        std::string data59 = sourceWks.cell("O22").value().getString();
        std::string data60 = sourceWks.cell("S22").value().getString();

        std::string data61 = sourceWks.cell("AR22").value().getString();
        std::string data62 = sourceWks.cell("AV22").value().getString();
        std::string data63 = sourceWks.cell("AZ22").value().getString();

        std::string data64 = sourceWks.cell("BY22").value().getString();
        std::string data65 = sourceWks.cell("CC22").value().getString();
        std::string data66 = sourceWks.cell("CG22").value().getString();
       
        // Для 7 предмета
        std::string data67 = sourceWks.cell("D23").value().getString();
        std::string data68 = sourceWks.cell("E23").value().getString();

        std::string data69 = sourceWks.cell("K23").value().getString();
        std::string data70 = sourceWks.cell("O23").value().getString();
        std::string data71 = sourceWks.cell("S23").value().getString();

        std::string data72 = sourceWks.cell("AR23").value().getString();
        std::string data73 = sourceWks.cell("AV23").value().getString();
        std::string data74 = sourceWks.cell("AZ23").value().getString();

        std::string data75 = sourceWks.cell("BY23").value().getString();
        std::string data76 = sourceWks.cell("CC23").value().getString();
        std::string data77 = sourceWks.cell("CG23").value().getString();

        // Для 8 предмета
        std::string data78 = sourceWks.cell("D24").value().getString();
        std::string data79 = sourceWks.cell("E24").value().getString();

        std::string data80 = sourceWks.cell("K24").value().getString();
        std::string data81 = sourceWks.cell("O24").value().getString();
        std::string data82 = sourceWks.cell("S24").value().getString();

        std::string data83 = sourceWks.cell("AR24").value().getString();
        std::string data84 = sourceWks.cell("AV24").value().getString();
        std::string data85 = sourceWks.cell("AZ24").value().getString();

        std::string data86 = sourceWks.cell("BY24").value().getString();
        std::string data87 = sourceWks.cell("CC24").value().getString();
        std::string data88 = sourceWks.cell("CG24").value().getString();

        // Для 9 предмета
        std::string data89 = sourceWks.cell("D25").value().getString();
        std::string data90 = sourceWks.cell("E25").value().getString();

        std::string data91 = sourceWks.cell("K25").value().getString();
        std::string data92 = sourceWks.cell("O25").value().getString();
        std::string data93 = sourceWks.cell("S25").value().getString();

        std::string data94 = sourceWks.cell("AR25").value().getString();
        std::string data95 = sourceWks.cell("AV25").value().getString();
        std::string data96 = sourceWks.cell("AZ25").value().getString();

        std::string data97 = sourceWks.cell("BY25").value().getString();
        std::string data98 = sourceWks.cell("CC25").value().getString();
        std::string data99 = sourceWks.cell("CG25").value().getString();

        // Для 10 предмета
        std::string data100 = sourceWks.cell("D26").value().getString();
        std::string data101 = sourceWks.cell("E26").value().getString();

        std::string data102 = sourceWks.cell("K26").value().getString();
        std::string data103 = sourceWks.cell("O26").value().getString();
        std::string data104 = sourceWks.cell("S26").value().getString();

        std::string data105 = sourceWks.cell("AR26").value().getString();
        std::string data106 = sourceWks.cell("AV26").value().getString();
        std::string data107 = sourceWks.cell("AZ26").value().getString();

        std::string data108 = sourceWks.cell("BY26").value().getString();
        std::string data109 = sourceWks.cell("CC26").value().getString();
        std::string data110 = sourceWks.cell("CG26").value().getString();

        // Для 11 предмета
        std::string data111 = sourceWks.cell("D27").value().getString();
        std::string data112 = sourceWks.cell("E27").value().getString();

        std::string data113 = sourceWks.cell("K27").value().getString();
        std::string data114 = sourceWks.cell("O27").value().getString();
        std::string data115 = sourceWks.cell("S27").value().getString();

        std::string data116 = sourceWks.cell("AR27").value().getString();
        std::string data117 = sourceWks.cell("AV27").value().getString();
        std::string data118 = sourceWks.cell("AZ27").value().getString();

        std::string data119 = sourceWks.cell("BY27").value().getString();
        std::string data120 = sourceWks.cell("CC27").value().getString();
        std::string data121 = sourceWks.cell("CG27").value().getString();

        // Для 12 предмета
        std::string data122 = sourceWks.cell("D28").value().getString();
        std::string data123 = sourceWks.cell("E28").value().getString();

        std::string data124 = sourceWks.cell("K28").value().getString();
        std::string data125 = sourceWks.cell("O28").value().getString();
        std::string data126 = sourceWks.cell("S28").value().getString();

        std::string data127 = sourceWks.cell("AR28").value().getString();
        std::string data128 = sourceWks.cell("AV28").value().getString();
        std::string data129 = sourceWks.cell("AZ28").value().getString();

        std::string data130 = sourceWks.cell("BY28").value().getString();
        std::string data131 = sourceWks.cell("CC28").value().getString();
        std::string data132 = sourceWks.cell("CG28").value().getString();

        // Для 13 предмета
        std::string data133 = sourceWks.cell("D28").value().getString();
        std::string data134 = sourceWks.cell("E28").value().getString();

        std::string data135 = sourceWks.cell("K28").value().getString();
        std::string data136 = sourceWks.cell("O28").value().getString();
        std::string data137 = sourceWks.cell("S28").value().getString();

        std::string data138 = sourceWks.cell("AR28").value().getString();
        std::string data139 = sourceWks.cell("AV28").value().getString();
        std::string data140 = sourceWks.cell("AZ28").value().getString();

        std::string data141 = sourceWks.cell("BY28").value().getString();
        std::string data142 = sourceWks.cell("CC28").value().getString();
        std::string data143 = sourceWks.cell("CG28").value().getString();

        // Для 14 предмета
        std::string data144 = sourceWks.cell("D29").value().getString();
        std::string data145 = sourceWks.cell("E29").value().getString();

        std::string data146 = sourceWks.cell("K29").value().getString();
        std::string data147 = sourceWks.cell("O29").value().getString();
        std::string data148 = sourceWks.cell("S29").value().getString();

        std::string data149 = sourceWks.cell("AR29").value().getString();
        std::string data150 = sourceWks.cell("AV29").value().getString();
        std::string data151 = sourceWks.cell("AZ29").value().getString();

        std::string data152 = sourceWks.cell("BY29").value().getString();
        std::string data153 = sourceWks.cell("CC29").value().getString();
        std::string data154 = sourceWks.cell("CG29").value().getString();

        // Для 15 предмета
        std::string data155 = sourceWks.cell("D30").value().getString();
        std::string data156 = sourceWks.cell("E30").value().getString();

        std::string data157 = sourceWks.cell("K30").value().getString();
        std::string data158 = sourceWks.cell("O30").value().getString();
        std::string data159 = sourceWks.cell("S30").value().getString();

        std::string data160 = sourceWks.cell("AR30").value().getString();
        std::string data161 = sourceWks.cell("AV30").value().getString();
        std::string data162 = sourceWks.cell("AZ30").value().getString();

        std::string data163 = sourceWks.cell("BY30").value().getString();
        std::string data164 = sourceWks.cell("CC30").value().getString();
        std::string data165 = sourceWks.cell("CG30").value().getString();

        // Для 16 предмета
        std::string data166 = sourceWks.cell("D31").value().getString();
        std::string data167 = sourceWks.cell("E31").value().getString();

        std::string data168 = sourceWks.cell("K31").value().getString();
        std::string data169 = sourceWks.cell("O31").value().getString();
        std::string data170 = sourceWks.cell("S31").value().getString();

        std::string data171 = sourceWks.cell("AR31").value().getString();
        std::string data172 = sourceWks.cell("AV31").value().getString();
        std::string data173 = sourceWks.cell("AZ31").value().getString();

        std::string data174 = sourceWks.cell("BY31").value().getString();
        std::string data175 = sourceWks.cell("CC31").value().getString();
        std::string data176 = sourceWks.cell("CG31").value().getString();

        // Для 17 предмета
        std::string data177 = sourceWks.cell("D32").value().getString();
        std::string data178 = sourceWks.cell("E32").value().getString();

        std::string data179 = sourceWks.cell("K32").value().getString();
        std::string data180 = sourceWks.cell("O32").value().getString();
        std::string data181 = sourceWks.cell("S32").value().getString();

        std::string data182 = sourceWks.cell("AR32").value().getString();
        std::string data183 = sourceWks.cell("AV32").value().getString();
        std::string data184 = sourceWks.cell("AZ32").value().getString();

        std::string data185 = sourceWks.cell("BY32").value().getString();
        std::string data186 = sourceWks.cell("CC32").value().getString();
        std::string data187 = sourceWks.cell("CG32").value().getString();

        // Открываем целевой файл
        OpenXLSX::XLDocument targetDoc;
        targetDoc.open("D:\\Test\\P0.xlsx"); // Путь к файлу (пустой бланк)
        auto targetWks = targetDoc.workbook().worksheet("Sheet1"); // Замените "Sheet1" на имя вашего листа

        // Записываем данные в целевой файл
        // Для 1 предмета
        targetWks.cell("A12").value() = data1;
        targetWks.cell("A9").value() = data2;

        targetWks.cell("C9").value() = toInt(data3);
        targetWks.cell("C10").value() = data5;
        targetWks.cell("C11").value() = data4;

        targetWks.cell("D9").value() = toInt(data6);
        targetWks.cell("D10").value() = toInt(data8);
        targetWks.cell("D11").value() = toInt(data7);

        targetWks.cell("E9").value() = toInt(data9);
        targetWks.cell("E10").value() = toInt(data11);
        targetWks.cell("E11").value() = toInt(data10);

        // Для 2 предмета
        targetWks.cell("A16").value() = data12;
        targetWks.cell("A13").value() = data13;

        targetWks.cell("C13").value() = data14;
        targetWks.cell("C14").value() = data16;
        targetWks.cell("C15").value() = data15;

        targetWks.cell("D13").value() = data17;
        targetWks.cell("D14").value() = data19;
        targetWks.cell("D15").value() = data18;

        targetWks.cell("E13").value() = data20;
        targetWks.cell("E14").value() = data22;
        targetWks.cell("E15").value() = data21;

        // Для 3 предмета
        targetWks.cell("A20").value() = data23;
        targetWks.cell("A17").value() = data24;

        targetWks.cell("C17").value() = data25;
        targetWks.cell("C18").value() = data27;
        targetWks.cell("C19").value() = data26;

        targetWks.cell("D17").value() = data28;
        targetWks.cell("D18").value() = data30;
        targetWks.cell("D19").value() = data29;

        targetWks.cell("E17").value() = data31;
        targetWks.cell("E18").value() = data33;
        targetWks.cell("E19").value() = data32;

        // Для 4 предмета
        targetWks.cell("A24").value() = data34;
        targetWks.cell("A21").value() = data35;

        targetWks.cell("C21").value() = data36;
        targetWks.cell("C22").value() = data38;
        targetWks.cell("C23").value() = data37;

        targetWks.cell("D21").value() = data39;
        targetWks.cell("D22").value() = data41;
        targetWks.cell("D23").value() = data40;

        targetWks.cell("E21").value() = data42;
        targetWks.cell("E22").value() = data44;
        targetWks.cell("E23").value() = data43;

        // Для 5 предмета
        targetWks.cell("A28").value() = data45;
        targetWks.cell("A25").value() = data46;

        targetWks.cell("C25").value() = data47;
        targetWks.cell("C26").value() = data49;
        targetWks.cell("C27").value() = data48;

        targetWks.cell("D25").value() = data50;
        targetWks.cell("D26").value() = data52;
        targetWks.cell("D27").value() = data51;

        targetWks.cell("E25").value() = data53;
        targetWks.cell("E26").value() = data55;
        targetWks.cell("E27").value() = data54;

        // Для 6 предмета
        targetWks.cell("A32").value() = data56;
        targetWks.cell("A29").value() = data57;

        targetWks.cell("C29").value() = data58;
        targetWks.cell("C30").value() = data60;
        targetWks.cell("C31").value() = data59;

        targetWks.cell("D29").value() = data61;
        targetWks.cell("D30").value() = data63;
        targetWks.cell("D31").value() = data62;

        targetWks.cell("E29").value() = data64;
        targetWks.cell("E30").value() = data66;
        targetWks.cell("E31").value() = data65;

        // Для 7 предмета
        targetWks.cell("A36").value() = data67;
        targetWks.cell("A33").value() = data68;

        targetWks.cell("C33").value() = data69;
        targetWks.cell("C34").value() = data71;
        targetWks.cell("C35").value() = data70;

        targetWks.cell("D33").value() = data72;
        targetWks.cell("D34").value() = data74;
        targetWks.cell("D35").value() = data73;

        targetWks.cell("E33").value() = data75;
        targetWks.cell("E34").value() = data77;
        targetWks.cell("E35").value() = data76;

        // Для 8 предмета
        targetWks.cell("A40").value() = data78;
        targetWks.cell("A37").value() = data79;

        targetWks.cell("C37").value() = data80;
        targetWks.cell("C38").value() = data82;
        targetWks.cell("C39").value() = data81;

        targetWks.cell("D37").value() = data83;
        targetWks.cell("D38").value() = data85;
        targetWks.cell("D39").value() = data84;

        targetWks.cell("E37").value() = data86;
        targetWks.cell("E38").value() = data88;
        targetWks.cell("E39").value() = data87;

        // Для 9 предмета
        targetWks.cell("A44").value() = data89;
        targetWks.cell("A41").value() = data90;

        targetWks.cell("C41").value() = data91;
        targetWks.cell("C42").value() = data93;
        targetWks.cell("C43").value() = data92;

        targetWks.cell("D41").value() = data94;
        targetWks.cell("D42").value() = data96;
        targetWks.cell("D43").value() = data95;

        targetWks.cell("E41").value() = data97;
        targetWks.cell("E42").value() = data99;
        targetWks.cell("E43").value() = data98;

        // Для 10 предмета
        targetWks.cell("A63").value() = data100;
        targetWks.cell("A60").value() = data101;

        targetWks.cell("C60").value() = data102;
        targetWks.cell("C61").value() = data104;
        targetWks.cell("C62").value() = data103;

        targetWks.cell("D60").value() = data105;
        targetWks.cell("D61").value() = data107;
        targetWks.cell("D62").value() = data106;

        targetWks.cell("E60").value() = data108;
        targetWks.cell("E61").value() = data110;
        targetWks.cell("E62").value() = data109;
       
        // Для 11 предмета
        targetWks.cell("A67").value() = data111;
        targetWks.cell("A64").value() = data112;

        targetWks.cell("C64").value() = data113;
        targetWks.cell("C65").value() = data115;
        targetWks.cell("C66").value() = data114;

        targetWks.cell("D64").value() = data116;
        targetWks.cell("D65").value() = data118;
        targetWks.cell("D66").value() = data117;

        targetWks.cell("E64").value() = data119;
        targetWks.cell("E65").value() = data121;
        targetWks.cell("E66").value() = data120;

        // Для 12 предмета
        targetWks.cell("A71").value() = data122;
        targetWks.cell("A68").value() = data123;

        targetWks.cell("C68").value() = data124;
        targetWks.cell("C69").value() = data126;
        targetWks.cell("C70").value() = data125;

        targetWks.cell("D68").value() = data127;
        targetWks.cell("D69").value() = data129;
        targetWks.cell("D70").value() = data128;

        targetWks.cell("E68").value() = data130;
        targetWks.cell("E69").value() = data132;
        targetWks.cell("E70").value() = data131;

        // Для 13 предмета
        targetWks.cell("A75").value() = data133;
        targetWks.cell("A72").value() = data134;

        targetWks.cell("C72").value() = data135;
        targetWks.cell("C73").value() = data137;
        targetWks.cell("C74").value() = data136;

        targetWks.cell("D72").value() = data138;
        targetWks.cell("D73").value() = data140;
        targetWks.cell("D74").value() = data139;

        targetWks.cell("E72").value() = data141;
        targetWks.cell("E73").value() = data143;
        targetWks.cell("E74").value() = data142;

        // Для 14 предмета
        targetWks.cell("A79").value() = data144;
        targetWks.cell("A76").value() = data145;

        targetWks.cell("C76").value() = data146;
        targetWks.cell("C77").value() = data148;
        targetWks.cell("C78").value() = data147;

        targetWks.cell("D76").value() = data149;
        targetWks.cell("D77").value() = data151;
        targetWks.cell("D78").value() = data150;

        targetWks.cell("E76").value() = data152;
        targetWks.cell("E77").value() = data154;
        targetWks.cell("E78").value() = data153;

        // Для 15 предмета
        targetWks.cell("A79").value() = data155;
        targetWks.cell("A76").value() = data156;

        targetWks.cell("C76").value() = data157;
        targetWks.cell("C77").value() = data159;
        targetWks.cell("C78").value() = data158;

        targetWks.cell("D76").value() = data160;
        targetWks.cell("D77").value() = data162;
        targetWks.cell("D78").value() = data161;

        targetWks.cell("E76").value() = data163;
        targetWks.cell("E77").value() = data165;
        targetWks.cell("E78").value() = data164;

        // Для 16 предмета
        targetWks.cell("A79").value() = data166;
        targetWks.cell("A76").value() = data167;

        targetWks.cell("C76").value() = data168;
        targetWks.cell("C77").value() = data170;
        targetWks.cell("C78").value() = data169;

        targetWks.cell("D76").value() = data171;
        targetWks.cell("D77").value() = data173;
        targetWks.cell("D78").value() = data172;

        targetWks.cell("E76").value() = data174;
        targetWks.cell("E77").value() = data176;
        targetWks.cell("E78").value() = data175;

        // Для 17 предмета
        targetWks.cell("A79").value() = data177;
        targetWks.cell("A76").value() = data178;

        targetWks.cell("C76").value() = data179;
        targetWks.cell("C77").value() = data181;
        targetWks.cell("C78").value() = data180;

        targetWks.cell("D76").value() = data182;
        targetWks.cell("D77").value() = data184;
        targetWks.cell("D78").value() = data183;

        targetWks.cell("E76").value() = data185;
        targetWks.cell("E77").value() = data187;
        targetWks.cell("E78").value() = data186;
        // Сохраняем изменения в целевом файле
        targetDoc.save();
    }
    catch (const std::exception& e) {
        std::cerr << "Exception caught : " << e.what() << std::endl;
    }
    return 0;
}