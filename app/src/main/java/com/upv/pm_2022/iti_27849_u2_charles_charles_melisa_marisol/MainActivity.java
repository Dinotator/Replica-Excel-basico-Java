package com.upv.pm_2022.iti_27849_u2_charles_charles_melisa_marisol;



import androidx.appcompat.app.AlertDialog;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.content.Context;
import android.content.DialogInterface;
import android.os.Bundle;
import android.os.Environment;
import android.view.LayoutInflater;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import com.obsez.android.lib.filechooser.ChooserDialog;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;

import java.util.HashMap;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {
    private static final int REQUEST_ID_READ_PERMISSION = 100;
    private static final int REQUEST_ID_WRITE_PERMISSION = 200;
    private Context CX;
    private final File filePath = new File(Environment.getExternalStorageDirectory().toString() + "/File.xlsx");

    private String Path;

    private String startingDir;
    private ArrayList<String> cellsName, valueOfCells;
    private ArrayList<TextView> cellsAsObjects;
    private ArrayList<Integer> cellNumber;

    public TextView a1, b1, c1, d1, e1, f1, g1,
            a2, b2, c2, d2, e2, f2, g2,
            a3, b3, c3, d3, e3, f3, g3,
            a4, b4, c4, d4, e4, f4, g4,
            a5, b5, c5, d5, e5, f5, g5,
            a6, b6, c6, d6, e6, f6, g6,
            a7, b7, c7, d7, e7, f7, g7,
            a8, b8, c8, d8, e8, f8, g8,
            a9, b9, c9, d9, e9, f9, g9,
            a10, b10, c10, d10, f10, e10, g10,
            a11, b11, c11, d11, f11, e11, g11,
            a12, b12, c12, d12, f12, e12, g12,
            a13, b13, c13, d13, f13, e13, g13,
            a14, b14, c14, d14, f14, e14, g14,
            a15, b15, c15, d15, f15, e15, g15,
            a16, b16, c16, d16, f16, e16, g16,
            a17, b17, c17, d17, f17, e17, g17,
            a18, b18, c18, d18, f18, e18, g18,
            a19, b19, c19, d19, f19, e19, g19,
            a20, b20, c20, d20, f20, e20, g20;


    public Button btnImport, btnExport;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        askPermissionOnly();
        cellsAsObjects = new ArrayList();
        cellsName = new ArrayList();
        valueOfCells = new ArrayList();
        cellNumber = new ArrayList();

        btnImport = findViewById(R.id.btn_import);
        btnExport = findViewById(R.id.btn_export);

        cellsName = new ArrayList() {{
            add("A1");
            add("B1");
            add("C1");
            add("D1");
            add("E1");
            add("F1");
            add("G1");

            add("A2");
            add("B2");
            add("C2");
            add("D2");
            add("E2");
            add("F2");
            add("G2");

            add("A3");
            add("B3");
            add("C3");
            add("D3");
            add("E3");
            add("F3");
            add("G3");

            add("A4");
            add("B4");
            add("C4");
            add("D4");
            add("E4");
            add("F4");
            add("G4");

            add("A5");
            add("B5");
            add("C5");
            add("D5");
            add("E5");
            add("F5");
            add("G5");

            add("A6");
            add("B6");
            add("C6");
            add("D6");
            add("E6");
            add("F6");
            add("G6");

            add("A7");
            add("B7");
            add("C7");
            add("D7");
            add("E7");
            add("F7");
            add("G7");

            add("A8");
            add("B8");
            add("C8");
            add("D8");
            add("E8");
            add("F8");
            add("G8");

            add("A9");
            add("B9");
            add("C9");
            add("D9");
            add("E9");
            add("F9");
            add("G9");

            add("A10");
            add("B10");
            add("C10");
            add("D10");
            add("E10");
            add("F10");
            add("G10");

            add("A11");
            add("B11");
            add("C11");
            add("D11");
            add("E11");
            add("F11");
            add("G11");

            add("A12");
            add("B12");
            add("C12");
            add("D12");
            add("E12");
            add("F12");
            add("G12");

            add("A13");
            add("B13");
            add("C13");
            add("D13");
            add("E13");
            add("F13");
            add("G13");

            add("A14");
            add("B14");
            add("C14");
            add("D14");
            add("E14");
            add("F14");
            add("G14");

            add("A15");
            add("B15");
            add("C15");
            add("D15");
            add("E15");
            add("F15");
            add("G15");

            add("A16");
            add("B16");
            add("C16");
            add("D16");
            add("E16");
            add("F16");
            add("G16");

            add("A17");
            add("B17");
            add("C17");
            add("D17");
            add("E17");
            add("F17");
            add("G17");

            add("A18");
            add("B18");
            add("C18");
            add("D18");
            add("E18");
            add("F18");
            add("G18");

            add("A19");
            add("B19");
            add("C19");
            add("D19");
            add("E19");
            add("F19");
            add("G19");

            add("A20");
            add("B20");
            add("C20");
            add("D20");
            add("E20");
            add("F20");
            add("G20");
        }};


        cellsAsObjects = new ArrayList() {{
            add(a1 = findViewById(R.id.a1));
            add(b1 = findViewById(R.id.b1));
            add(c1 = findViewById(R.id.c1));
            add(d1 = findViewById(R.id.d1));
            add(e1 = findViewById(R.id.e1));
            add(f1 = findViewById(R.id.f1));
            add(g1 = findViewById(R.id.g1));

            add(a2 = findViewById(R.id.a2));
            add(b2 = findViewById(R.id.b2));
            add(c2 = findViewById(R.id.c2));
            add(d2 = findViewById(R.id.d2));
            add(e2 = findViewById(R.id.e2));
            add(f2 = findViewById(R.id.f2));
            add(g2 = findViewById(R.id.g2));

            add(a3 = findViewById(R.id.a3));
            add(b3 = findViewById(R.id.b3));
            add(c3 = findViewById(R.id.c3));
            add(d3 = findViewById(R.id.d3));
            add(e3 = findViewById(R.id.e3));
            add(f3 = findViewById(R.id.f3));
            add(g3 = findViewById(R.id.g3));

            add(a4 = findViewById(R.id.a4));
            add(b4 = findViewById(R.id.b4));
            add(c4 = findViewById(R.id.c4));
            add(d4 = findViewById(R.id.d4));
            add(e4 = findViewById(R.id.e4));
            add(f4 = findViewById(R.id.f4));
            add(g4 = findViewById(R.id.g4));

            add(a5 = findViewById(R.id.a5));
            add(b5 = findViewById(R.id.b5));
            add(c5 = findViewById(R.id.c5));
            add(d5 = findViewById(R.id.d5));
            add(e5 = findViewById(R.id.e5));
            add(f5 = findViewById(R.id.f5));
            add(g5 = findViewById(R.id.g5));

            add(a6 = findViewById(R.id.a6));
            add(b6 = findViewById(R.id.b6));
            add(c6 = findViewById(R.id.c6));
            add(d6 = findViewById(R.id.d6));
            add(e6 = findViewById(R.id.e6));
            add(f6 = findViewById(R.id.f6));
            add(g6 = findViewById(R.id.g6));

            add(a7 = findViewById(R.id.a7));
            add(b7 = findViewById(R.id.b7));
            add(c7 = findViewById(R.id.c7));
            add(d7 = findViewById(R.id.d7));
            add(e7 = findViewById(R.id.e7));
            add(f7 = findViewById(R.id.f7));
            add(g7 = findViewById(R.id.g7));

            add(a8 = findViewById(R.id.a8));
            add(b8 = findViewById(R.id.b8));
            add(c8 = findViewById(R.id.c8));
            add(d8 = findViewById(R.id.d8));
            add(e8 = findViewById(R.id.e8));
            add(f8 = findViewById(R.id.f8));
            add(g8 = findViewById(R.id.g8));

            add(a9 = findViewById(R.id.a9));
            add(b9 = findViewById(R.id.b9));
            add(c9 = findViewById(R.id.c9));
            add(d9 = findViewById(R.id.d9));
            add(e9 = findViewById(R.id.e9));
            add(f9 = findViewById(R.id.f9));
            add(g9 = findViewById(R.id.g9));

            add(a10 = findViewById(R.id.a10));
            add(b10 = findViewById(R.id.b10));
            add(c10 = findViewById(R.id.c10));
            add(d10 = findViewById(R.id.d10));
            add(e10 = findViewById(R.id.e10));
            add(f10 = findViewById(R.id.f10));
            add(g10 = findViewById(R.id.g10));

            add(a11 = findViewById(R.id.a11));
            add(b11 = findViewById(R.id.b11));
            add(c11 = findViewById(R.id.c11));
            add(d11 = findViewById(R.id.d11));
            add(e11 = findViewById(R.id.e11));
            add(f11 = findViewById(R.id.f11));
            add(g11 = findViewById(R.id.g11));

            add(a12 = findViewById(R.id.a12));
            add(b12 = findViewById(R.id.b12));
            add(c12 = findViewById(R.id.c12));
            add(d12 = findViewById(R.id.d12));
            add(e12 = findViewById(R.id.e12));
            add(f12 = findViewById(R.id.f12));
            add(g12 = findViewById(R.id.g12));

            add(a13 = findViewById(R.id.a13));
            add(b13 = findViewById(R.id.b13));
            add(c13 = findViewById(R.id.c13));
            add(d13 = findViewById(R.id.d13));
            add(e13 = findViewById(R.id.e13));
            add(f13 = findViewById(R.id.f13));
            add(g13 = findViewById(R.id.g13));

            add(a14 = findViewById(R.id.a14));
            add(b14 = findViewById(R.id.b14));
            add(c14 = findViewById(R.id.c14));
            add(d14 = findViewById(R.id.d14));
            add(e14 = findViewById(R.id.e14));
            add(f14 = findViewById(R.id.f14));
            add(g14 = findViewById(R.id.g14));

            add(a15 = findViewById(R.id.a15));
            add(b15 = findViewById(R.id.b15));
            add(c15 = findViewById(R.id.c15));
            add(d15 = findViewById(R.id.d15));
            add(e15 = findViewById(R.id.e15));
            add(f15 = findViewById(R.id.f15));
            add(g15 = findViewById(R.id.g15));

            add(a16 = findViewById(R.id.a16));
            add(b16 = findViewById(R.id.b16));
            add(c16 = findViewById(R.id.c16));
            add(d16 = findViewById(R.id.d16));
            add(e16 = findViewById(R.id.e16));
            add(f16 = findViewById(R.id.f16));
            add(g16 = findViewById(R.id.g16));

            add(a17 = findViewById(R.id.a17));
            add(b17 = findViewById(R.id.b17));
            add(c17 = findViewById(R.id.c17));
            add(d17 = findViewById(R.id.d17));
            add(e17 = findViewById(R.id.e17));
            add(f17 = findViewById(R.id.f17));
            add(g17 = findViewById(R.id.g17));

            add(a18 = findViewById(R.id.a18));
            add(b18 = findViewById(R.id.b18));
            add(c18 = findViewById(R.id.c18));
            add(d18 = findViewById(R.id.d18));
            add(e18 = findViewById(R.id.e18));
            add(f18 = findViewById(R.id.f18));
            add(g18 = findViewById(R.id.g18));

            add(a19 = findViewById(R.id.a19));
            add(b19 = findViewById(R.id.b19));
            add(c19 = findViewById(R.id.c19));
            add(d19 = findViewById(R.id.d19));
            add(e19 = findViewById(R.id.e19));
            add(f19 = findViewById(R.id.f19));
            add(g19 = findViewById(R.id.g19));

            add(a20 = findViewById(R.id.a20));
            add(b20 = findViewById(R.id.b20));
            add(c20 = findViewById(R.id.c20));
            add(d20 = findViewById(R.id.d20));
            add(e20 = findViewById(R.id.e20));
            add(f20 = findViewById(R.id.f20));
            add(g20 = findViewById(R.id.g20));

        }};
        for (int i = 0; i < cellsAsObjects.size(); i++) {
            cellNumber.add(i);
        }

        for (int i = 0; i < cellsAsObjects.size(); i++) {
            cellsAsObjects.get(i).setOnClickListener(this);
        }


        for (int i = 0; i < cellsAsObjects.size(); i++) {
            valueOfCells.add(cellsAsObjects.get(i).getText().toString());
        }
        btnExport.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                try {
                    XSSFWorkbook XSSFWorkbook = new XSSFWorkbook();

                    XSSFSheet Sheet = XSSFWorkbook.createSheet("Archivo Generado");

                    XSSFRow row = Sheet.createRow(0);
                    XSSFCell cell = row.createCell(0);

                    int aux = 0;

                    for (int i = 0; i < 20; i++) {
                        row = Sheet.createRow(i);
                        for (int j = 0; j < 7; j++) {
                            cell = row.createCell(j);
                            cell.setCellValue(valueOfCells.get(aux));
                            aux++;
                        }
                    }

                    try {
                        if (!filePath.exists()) {
                            filePath.createNewFile();
                        }

                        OutputStream fileOutputStream = new FileOutputStream(filePath);
                        XSSFWorkbook.write(fileOutputStream);

                        if (fileOutputStream != null) {
                            fileOutputStream.flush();
                            fileOutputStream.close();
                        }
                    } catch (Exception e) {

                    }

                    Toast.makeText(getApplicationContext(), "export successful", Toast.LENGTH_SHORT).show();
                } catch (Exception e) {
                    Toast.makeText(getApplicationContext(), "export error" + e, Toast.LENGTH_SHORT).show();

                }
            }

        });

        btnImport.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                new ChooserDialog(MainActivity.this)
                        .withFilter(false, false, "xlsx", "XLSX")
                        .withStartFile(startingDir)
                        .withResources(R.string.app_name,R.string.yes_button,R.string.no_button)
                        .withChosenListener(new ChooserDialog.Result() {
                            @Override
                            public void onChoosePath(String path, File pathFile) {
                                Path = path;
                                readExcel();
                            }
                        })
                        .build()
                        .show();
            }

        });

    }



    @Override
    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.a1:
                putTextInCell(cellsName.get(0), cellNumber.get(0));
                break;
            case R.id.b1:
                putTextInCell(cellsName.get(1), cellNumber.get(1));
                break;
            case R.id.c1:
                putTextInCell(cellsName.get(2), cellNumber.get(2));
                break;
            case R.id.d1:
                putTextInCell(cellsName.get(3), cellNumber.get(3));
                break;
            case R.id.e1:
                putTextInCell(cellsName.get(4), cellNumber.get(4));
                break;
            case R.id.f1:
                putTextInCell(cellsName.get(5), cellNumber.get(5));
                break;
            case R.id.g1:
                putTextInCell(cellsName.get(6), cellNumber.get(6));
                break;

            case R.id.a2:
                putTextInCell(cellsName.get(7), cellNumber.get(7));
                break;
            case R.id.b2:
                putTextInCell(cellsName.get(8), cellNumber.get(8));
                break;
            case R.id.c2:
                putTextInCell(cellsName.get(9), cellNumber.get(9));
                break;
            case R.id.d2:
                putTextInCell(cellsName.get(10), cellNumber.get(10));
                break;
            case R.id.e2:
                putTextInCell(cellsName.get(11), cellNumber.get(11));
                break;
            case R.id.f2:
                putTextInCell(cellsName.get(12), cellNumber.get(12));
                break;
            case R.id.g2:
                putTextInCell(cellsName.get(13), cellNumber.get(13));
                break;

            case R.id.a3:
                putTextInCell(cellsName.get(14), cellNumber.get(14));
                break;
            case R.id.b3:
                putTextInCell(cellsName.get(15), cellNumber.get(15));
                break;
            case R.id.c3:
                putTextInCell(cellsName.get(16), cellNumber.get(16));
                break;
            case R.id.d3:
                putTextInCell(cellsName.get(17), cellNumber.get(17));
                break;
            case R.id.e3:
                putTextInCell(cellsName.get(18), cellNumber.get(18));
                break;
            case R.id.f3:
                putTextInCell(cellsName.get(19), cellNumber.get(19));
                break;
            case R.id.g3:
                putTextInCell(cellsName.get(20), cellNumber.get(20));
                break;

            case R.id.a4:
                putTextInCell(cellsName.get(21), cellNumber.get(21));
                break;
            case R.id.b4:
                putTextInCell(cellsName.get(22), cellNumber.get(22));
                break;
            case R.id.c4:
                putTextInCell(cellsName.get(23), cellNumber.get(23));
                break;
            case R.id.d4:
                putTextInCell(cellsName.get(24), cellNumber.get(24));
                break;
            case R.id.e4:
                putTextInCell(cellsName.get(25), cellNumber.get(25));
                break;
            case R.id.f4:
                putTextInCell(cellsName.get(26), cellNumber.get(26));
                break;
            case R.id.g4:
                putTextInCell(cellsName.get(27), cellNumber.get(27));
                break;

            case R.id.a5:
                putTextInCell(cellsName.get(28), cellNumber.get(28));
                break;
            case R.id.b5:
                putTextInCell(cellsName.get(29), cellNumber.get(29));
                break;
            case R.id.c5:
                putTextInCell(cellsName.get(30), cellNumber.get(30));
                break;
            case R.id.d5:
                putTextInCell(cellsName.get(31), cellNumber.get(31));
                break;
            case R.id.e5:
                putTextInCell(cellsName.get(32), cellNumber.get(32));
                break;
            case R.id.f5:
                putTextInCell(cellsName.get(33), cellNumber.get(33));
                break;
            case R.id.g5:
                putTextInCell(cellsName.get(34), cellNumber.get(34));
                break;

            case R.id.a6:
                putTextInCell(cellsName.get(35), cellNumber.get(35));
                break;
            case R.id.b6:
                putTextInCell(cellsName.get(36), cellNumber.get(36));
                break;
            case R.id.c6:
                putTextInCell(cellsName.get(37), cellNumber.get(37));
                break;
            case R.id.d6:
                putTextInCell(cellsName.get(38), cellNumber.get(38));
                break;
            case R.id.e6:
                putTextInCell(cellsName.get(39), cellNumber.get(39));
                break;
            case R.id.f6:
                putTextInCell(cellsName.get(40), cellNumber.get(40));
                break;
            case R.id.g6:
                putTextInCell(cellsName.get(41), cellNumber.get(41));
                break;

            case R.id.a7:
                putTextInCell(cellsName.get(42), cellNumber.get(42));
                break;
            case R.id.b7:
                putTextInCell(cellsName.get(43), cellNumber.get(43));
                break;
            case R.id.c7:
                putTextInCell(cellsName.get(44), cellNumber.get(44));
                break;
            case R.id.d7:
                putTextInCell(cellsName.get(45), cellNumber.get(45));
                break;
            case R.id.e7:
                putTextInCell(cellsName.get(46), cellNumber.get(46));
                break;
            case R.id.f7:
                putTextInCell(cellsName.get(47), cellNumber.get(47));
                break;
            case R.id.g7:
                putTextInCell(cellsName.get(48), cellNumber.get(48));
                break;

            case R.id.a8:
                putTextInCell(cellsName.get(49), cellNumber.get(49));
                break;
            case R.id.b8:
                putTextInCell(cellsName.get(50), cellNumber.get(50));
                break;
            case R.id.c8:
                putTextInCell(cellsName.get(51), cellNumber.get(51));
                break;
            case R.id.d8:
                putTextInCell(cellsName.get(52), cellNumber.get(52));
                break;
            case R.id.e8:
                putTextInCell(cellsName.get(53), cellNumber.get(53));
                break;
            case R.id.f8:
                putTextInCell(cellsName.get(54), cellNumber.get(54));
                break;
            case R.id.g8:
                putTextInCell(cellsName.get(55), cellNumber.get(55));
                break;

            case R.id.a9:
                putTextInCell(cellsName.get(56), cellNumber.get(56));
                break;
            case R.id.b9:
                putTextInCell(cellsName.get(57), cellNumber.get(57));
                break;
            case R.id.c9:
                putTextInCell(cellsName.get(58), cellNumber.get(58));
                break;
            case R.id.d9:
                putTextInCell(cellsName.get(59), cellNumber.get(59));
                break;
            case R.id.e9:
                putTextInCell(cellsName.get(60), cellNumber.get(60));
                break;
            case R.id.f9:
                putTextInCell(cellsName.get(61), cellNumber.get(61));
                break;
            case R.id.g9:
                putTextInCell(cellsName.get(62), cellNumber.get(62));
                break;

            case R.id.a10:
                putTextInCell(cellsName.get(63), cellNumber.get(63));
                break;
            case R.id.b10:
                putTextInCell(cellsName.get(64), cellNumber.get(64));
                break;
            case R.id.c10:
                putTextInCell(cellsName.get(65), cellNumber.get(65));
                break;
            case R.id.d10:
                putTextInCell(cellsName.get(66), cellNumber.get(66));
                break;
            case R.id.e10:
                putTextInCell(cellsName.get(67), cellNumber.get(67));
                break;
            case R.id.f10:
                putTextInCell(cellsName.get(68), cellNumber.get(68));
                break;
            case R.id.g10:
                putTextInCell(cellsName.get(69), cellNumber.get(69));
                break;

            case R.id.a11:
                putTextInCell(cellsName.get(70), cellNumber.get(70));
                break;
            case R.id.b11:
                putTextInCell(cellsName.get(71), cellNumber.get(71));
                break;
            case R.id.c11:
                putTextInCell(cellsName.get(72), cellNumber.get(72));
                break;
            case R.id.d11:
                putTextInCell(cellsName.get(73), cellNumber.get(73));
                break;
            case R.id.e11:
                putTextInCell(cellsName.get(74), cellNumber.get(74));
                break;
            case R.id.f11:
                putTextInCell(cellsName.get(75), cellNumber.get(75));
                break;
            case R.id.g11:
                putTextInCell(cellsName.get(76), cellNumber.get(76));
                break;

            case R.id.a12:
                putTextInCell(cellsName.get(77), cellNumber.get(77));
                break;
            case R.id.b12:
                putTextInCell(cellsName.get(78), cellNumber.get(78));
                break;
            case R.id.c12:
                putTextInCell(cellsName.get(79), cellNumber.get(79));
                break;
            case R.id.d12:
                putTextInCell(cellsName.get(80), cellNumber.get(80));
                break;
            case R.id.e12:
                putTextInCell(cellsName.get(81), cellNumber.get(81));
                break;
            case R.id.f12:
                putTextInCell(cellsName.get(82), cellNumber.get(82));
                break;
            case R.id.g12:
                putTextInCell(cellsName.get(83), cellNumber.get(83));
                break;

            case R.id.a13:
                putTextInCell(cellsName.get(84), cellNumber.get(84));
                break;
            case R.id.b13:
                putTextInCell(cellsName.get(85), cellNumber.get(85));
                break;
            case R.id.c13:
                putTextInCell(cellsName.get(86), cellNumber.get(86));
                break;
            case R.id.d13:
                putTextInCell(cellsName.get(87), cellNumber.get(87));
                break;
            case R.id.e13:
                putTextInCell(cellsName.get(88), cellNumber.get(88));
                break;
            case R.id.f13:
                putTextInCell(cellsName.get(89), cellNumber.get(89));
                break;
            case R.id.g13:
                putTextInCell(cellsName.get(90), cellNumber.get(90));
                break;

            case R.id.a14:
                putTextInCell(cellsName.get(91), cellNumber.get(91));
                break;
            case R.id.b14:
                putTextInCell(cellsName.get(92), cellNumber.get(92));
                break;
            case R.id.c14:
                putTextInCell(cellsName.get(93), cellNumber.get(93));
                break;
            case R.id.d14:
                putTextInCell(cellsName.get(94), cellNumber.get(94));
                break;
            case R.id.e14:
                putTextInCell(cellsName.get(95), cellNumber.get(95));
                break;
            case R.id.f14:
                putTextInCell(cellsName.get(96), cellNumber.get(96));
                break;
            case R.id.g14:
                putTextInCell(cellsName.get(97), cellNumber.get(97));
                break;

            case R.id.a15:
                putTextInCell(cellsName.get(98), cellNumber.get(98));
                break;
            case R.id.b15:
                putTextInCell(cellsName.get(99), cellNumber.get(99));
                break;
            case R.id.c15:
                putTextInCell(cellsName.get(100), cellNumber.get(100));
                break;
            case R.id.d15:
                putTextInCell(cellsName.get(101), cellNumber.get(101));
                break;
            case R.id.e15:
                putTextInCell(cellsName.get(102), cellNumber.get(102));
                break;
            case R.id.f15:
                putTextInCell(cellsName.get(103), cellNumber.get(103));
                break;
            case R.id.g15:
                putTextInCell(cellsName.get(104), cellNumber.get(104));
                break;

            case R.id.a16:
                putTextInCell(cellsName.get(105), cellNumber.get(105));
                break;
            case R.id.b16:
                putTextInCell(cellsName.get(106), cellNumber.get(106));
                break;
            case R.id.c16:
                putTextInCell(cellsName.get(107), cellNumber.get(107));
                break;
            case R.id.d16:
                putTextInCell(cellsName.get(108), cellNumber.get(108));
                break;
            case R.id.e16:
                putTextInCell(cellsName.get(109), cellNumber.get(109));
                break;
            case R.id.f16:
                putTextInCell(cellsName.get(110), cellNumber.get(110));
                break;
            case R.id.g16:
                putTextInCell(cellsName.get(111), cellNumber.get(111));
                break;

            case R.id.a17:
                putTextInCell(cellsName.get(112), cellNumber.get(112));
                break;
            case R.id.b17:
                putTextInCell(cellsName.get(113), cellNumber.get(113));
                break;
            case R.id.c17:
                putTextInCell(cellsName.get(114), cellNumber.get(114));
                break;
            case R.id.d17:
                putTextInCell(cellsName.get(115), cellNumber.get(115));
                break;
            case R.id.e17:
                putTextInCell(cellsName.get(116), cellNumber.get(116));
                break;
            case R.id.f17:
                putTextInCell(cellsName.get(117), cellNumber.get(117));
                break;
            case R.id.g17:
                putTextInCell(cellsName.get(118), cellNumber.get(118));
                break;

            case R.id.a18:
                putTextInCell(cellsName.get(119), cellNumber.get(119));
                break;
            case R.id.b18:
                putTextInCell(cellsName.get(120), cellNumber.get(120));
                break;
            case R.id.c18:
                putTextInCell(cellsName.get(121), cellNumber.get(121));
                break;
            case R.id.d18:
                putTextInCell(cellsName.get(122), cellNumber.get(122));
                break;
            case R.id.e18:
                putTextInCell(cellsName.get(123), cellNumber.get(123));
                break;
            case R.id.f18:
                putTextInCell(cellsName.get(124), cellNumber.get(124));
                break;
            case R.id.g18:
                putTextInCell(cellsName.get(125), cellNumber.get(125));
                break;

            case R.id.a19:
                putTextInCell(cellsName.get(126), cellNumber.get(126));
                break;
            case R.id.b19:
                putTextInCell(cellsName.get(127), cellNumber.get(127));
                break;
            case R.id.c19:
                putTextInCell(cellsName.get(128), cellNumber.get(128));
                break;
            case R.id.d19:
                putTextInCell(cellsName.get(129), cellNumber.get(129));
                break;
            case R.id.e19:
                putTextInCell(cellsName.get(130), cellNumber.get(130));
                break;
            case R.id.f19:
                putTextInCell(cellsName.get(131), cellNumber.get(131));
                break;
            case R.id.g19:
                putTextInCell(cellsName.get(132), cellNumber.get(132));
                break;

            case R.id.a20:
                putTextInCell(cellsName.get(133), cellNumber.get(133));
                break;
            case R.id.b20:
                putTextInCell(cellsName.get(134), cellNumber.get(134));
                break;
            case R.id.c20:
                putTextInCell(cellsName.get(135), cellNumber.get(135));
                break;
            case R.id.d20:
                putTextInCell(cellsName.get(136), cellNumber.get(136));
                break;
            case R.id.e20:
                putTextInCell(cellsName.get(137), cellNumber.get(137));
                break;
            case R.id.f20:
                putTextInCell(cellsName.get(138), cellNumber.get(138));
                break;
            case R.id.g20:
                putTextInCell(cellsName.get(139), cellNumber.get(139));
                break;

        }


    }


    public void putTextInCell(String cellName, Integer numberOfCell) {
        AlertDialog.Builder builder = new AlertDialog.Builder(MainActivity.this);
        View v = LayoutInflater.from(MainActivity.this).inflate(R.layout.item_dialog, null, false);
        builder.setTitle(cellName);
        final EditText textValue = v.findViewById(R.id.textValue);

        textValue.setText(valueOfCells.get(numberOfCell));

        builder.setView(v);

        builder.setPositiveButton("Put Text", new DialogInterface.OnClickListener() {
            @Override
            public void onClick(DialogInterface dialog, int which) {
                valueOfCells.set(numberOfCell, selectFunction(textValue.getText().toString()));
                cellsAsObjects.get(numberOfCell).setText(selectFunction(textValue.getText().toString()));
            }
        });

        builder.setNegativeButton("Cancel", new DialogInterface.OnClickListener() {
            @Override
            public void onClick(DialogInterface dialog, int which) {
                dialog.dismiss();
            }
        });

        builder.show();
    }

    public String selectFunction(String data) {
        try {
            String[] cutWord = data.split("\\(");
            String keyWord = cutWord[0].toUpperCase();
            String valueCut = cutWord[1];
            String[] keyData = valueCut.split("\\)");
            String dataTraveler = keyData[0];

            if (keyWord.equals("SUMA") || keyWord.equals("SUM")) {
                return plus(dataTraveler);
            } else if (keyWord.equals("PROMEDIO") || keyWord.equals("AVERAGE")) {
                return average(dataTraveler);

            } else if (keyWord.equals("MAX")) {
                return maximum(dataTraveler);

            } else if (keyWord.equals("MIN")) {
                return minimum(dataTraveler);

            } else if (keyWord.equals("MODA") || keyWord.equals("MODE")) {
                return mode(dataTraveler);

            } else if (keyWord.equals("CONCAT") || keyWord.equals("CONCATENATE") || keyWord.equals("CONCATENAR")) {
                return concatenate(dataTraveler);
            }

        } catch (Exception e) {

        }
        return data;
    }

    public String plus(String valor) {
        try {
            String sentResult = "0";
            String[] cutWord = valor.split("\\(");
            String[] keyWord = cutWord[0].split(":");
            double[] containNumbers = new double[keyWord.length];
            double amount = 0.0;
            for (int i = 0; i < keyWord.length; i++) {
                try {
                    containNumbers[i] = Double.parseDouble(keyWord[i]);
                } catch (Exception e) {
                    for (int j = 0; j < cellsName.size(); j++) {
                        if (keyWord[i].equals(cellsName.get(j))) {
                            try {
                                containNumbers[i] = Double.parseDouble(cellsAsObjects.get(j).getText().toString());
                            } catch (Exception a) {
                                containNumbers[i] = 0.0;
                            }

                        }
                    }
                }
                amount = containNumbers[i] + amount;
            }
            sentResult = String.format("%.1f", amount);
            return sentResult;
        } catch (Exception e) {
            return "ERROR";
        }
    }

    public String average(String valor) {
        try {
            String sentResult = "0";
            String[] cutWord = valor.split("\\(");
            String[] keyWord = cutWord[0].split(":");
            double[] containNumbers = new double[keyWord.length];
            double amount = 0.0;
            for (int i = 0; i < keyWord.length; i++) {
                try {
                    containNumbers[i] = Double.parseDouble(keyWord[i]);
                } catch (Exception e) {
                    for (int j = 0; j < cellsName.size(); j++) {
                        if (keyWord[i].equals(cellsName.get(j))) {
                            try {
                                containNumbers[i] = Double.parseDouble(cellsAsObjects.get(j).getText().toString());
                            } catch (Exception a) {
                                containNumbers[i] = 0.0;
                            }

                        }
                    }
                }
                amount = containNumbers[i] + amount;
            }
            amount = amount / containNumbers.length;
            sentResult = String.format("%.1f", amount);
            return sentResult;
        } catch (Exception e) {
            return "ERROR";
        }

    }

    public String maximum(String valor) {
        try {
            String sentResult = "0";
            String[] cutWord = valor.split("\\(");
            String[] keyWord = cutWord[0].split(":");
            double[] containNumbers = new double[keyWord.length];
            double maxValue = 0.0;
            for (int i = 0; i < keyWord.length; i++) {
                try {
                    containNumbers[i] = Double.parseDouble(keyWord[i]);
                } catch (Exception e) {
                    for (int j = 0; j < cellsName.size(); j++) {
                        if (keyWord[i].equals(cellsName.get(j))) {
                            try {
                                containNumbers[i] = Double.parseDouble(cellsAsObjects.get(j).getText().toString());
                            } catch (Exception a) {
                                containNumbers[i] = 0.0;
                            }

                        }
                    }
                }

                for (int counter = 0; counter < containNumbers.length; counter++) {
                    if (containNumbers[counter] > maxValue) {
                        maxValue = containNumbers[counter];
                    }
                }
            }
            sentResult = String.format("%.1f", maxValue);
            return sentResult;
        } catch (Exception e) {
            return "ERROR";
        }

    }

    public String minimum(String valor) {
        try {
            String sentResult = "0";
            String[] cutWord = valor.split("\\(");
            String[] keyWord = cutWord[0].split(":");
            double[] containNumbers = new double[keyWord.length];

            for (int i = 0; i < keyWord.length; i++) {
                try {
                    containNumbers[i] = Double.parseDouble(keyWord[i]);
                } catch (Exception e) {
                    for (int j = 0; j < cellsName.size(); j++) {
                        if (keyWord[i].equals(cellsName.get(j))) {
                            try {
                                containNumbers[i] = Double.parseDouble(cellsAsObjects.get(j).getText().toString());
                            } catch (Exception a) {
                                containNumbers[i] = 0.0;
                            }

                        }
                    }
                }

            }
            double minValue = containNumbers[0];
            for (int counter = 0; counter < containNumbers.length; counter++) {

                if (containNumbers[counter] < minValue) {
                    minValue = containNumbers[counter];
                } else {

                }
            }

            sentResult = String.format("%.1f", minValue);
            return sentResult;
        } catch (Exception e) {
            return "ERROR";
        }
    }

    public String mode(String valor) {
        try {
            HashMap<Double, Double> hm = new HashMap<Double, Double>();
            double max = 1;
            double temp = 0;
            String sentResult = "0";
            String[] cutWord = valor.split("\\(");
            String[] keyWord = cutWord[0].split(":");
            double[] containNumbers = new double[keyWord.length];

            for (int i = 0; i < keyWord.length; i++) {
                try {
                    containNumbers[i] = Double.parseDouble(keyWord[i]);
                } catch (Exception e) {
                    for (int j = 0; j < cellsName.size(); j++) {
                        if (keyWord[i].equals(cellsName.get(j))) {
                            try {
                                containNumbers[i] = Double.parseDouble(cellsAsObjects.get(j).getText().toString());
                            } catch (Exception a) {
                                containNumbers[i] = 0.0;
                            }

                        }
                    }
                }

            }
            for (int i = 0; i < containNumbers.length; i++) {

                if (hm.get(containNumbers[i]) != null) {

                    double count = hm.get(containNumbers[i]);
                    count++;
                    hm.put(containNumbers[i], count);

                    if (count > max) {
                        max = count;
                        temp = containNumbers[i];
                    }
                } else
                    hm.put(containNumbers[i], 1.0);
            }

            sentResult = String.format("%.1f", temp);
            return sentResult;
        } catch (Exception e) {
            return "ERROR";
        }
    }

    public String concatenate(String valor) {
        try {

            String[] cutWord = valor.split("\\)");
            String[] keyWord = cutWord[0].split(":");
            String[] containWords = new String[keyWord.length];
            String group = "";


            for (int i = 0; i < keyWord.length; i++) {

                for (int j = 0; j < cellsName.size(); j++) {
                    if (keyWord[i].equals(cellsName.get(j))) {
                        containWords[i] = cellsAsObjects.get(j).getText().toString();
                    }

                }
                group = group + containWords[i];
            }
            return group;
        } catch (Exception e) {
            return "ERROR";
        }
    }

    private void readExcel(){
        InputStream myInput;
        try {
            myInput =  new FileInputStream(Path);
            Workbook workbook = new XSSFWorkbook (myInput);
            Sheet sheet = workbook.getSheetAt(0);
            int aux = 0;

            for (int i = 0;i < 20; i++) {
                Row row = sheet.getRow(i);
                for (int j = 0;j < 7; j++) {
                    Cell cell = row.getCell(j);
                    if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                        cellsAsObjects.get(aux).setText("");
                        valueOfCells.set(aux,"");
                    } else {

                        cellsAsObjects.get(aux).setText(selectFunction(cell.toString()));
                        valueOfCells.set(aux,selectFunction(cell.toString()));
                    }
                    aux++;
                }
            }
            Toast.makeText(getApplicationContext(), "File imported successfuly", Toast.LENGTH_SHORT).show();

        } catch (Exception e){

        }
    }

    private void askPermissionOnly() {
        this.askPermission(REQUEST_ID_WRITE_PERMISSION,
                android.Manifest.permission.WRITE_EXTERNAL_STORAGE);

        this.askPermission(REQUEST_ID_READ_PERMISSION,
                android.Manifest.permission.READ_EXTERNAL_STORAGE);

    }


    // With Android Level >= 23, you have to ask the user
    // for permission with device (For example read/write data on the device).
    private boolean askPermission(int requestId, String permissionName) {
        if (android.os.Build.VERSION.SDK_INT >= 23) {

            // Check if we have permission
            int permission = ActivityCompat.checkSelfPermission(this, permissionName);


            if (permission != android.content.pm.PackageManager.PERMISSION_GRANTED) {
                // If don't have permission so prompt the user.
                this.requestPermissions(
                        new String[]{permissionName},
                        requestId
                );
                return false;
            }
        }
        return true;
    }

    // When you have the request results
    @Override
    public void onRequestPermissionsResult(int requestCode,
                                           String permissions[], int[] grantResults) {

        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        //
        // Note: If request is cancelled, the result arrays are empty.
        if (grantResults.length > 0) {
            switch (requestCode) {
                case REQUEST_ID_READ_PERMISSION: {
                    if (grantResults[0] == android.content.pm.PackageManager.PERMISSION_GRANTED) {
                        Toast.makeText(getApplicationContext(), "Permission Lectura Concedido!", Toast.LENGTH_SHORT).show();
                    }
                }
                case REQUEST_ID_WRITE_PERMISSION: {
                    if (grantResults[0] == android.content.pm.PackageManager.PERMISSION_GRANTED) {
                        Toast.makeText(getApplicationContext(), "Permission Escritura Concedido!", Toast.LENGTH_SHORT).show();
                    }
                }
            }
        } else {
            Toast.makeText(getApplicationContext(), "Permission Cancelled!", Toast.LENGTH_SHORT).show();
        }
    }


}



