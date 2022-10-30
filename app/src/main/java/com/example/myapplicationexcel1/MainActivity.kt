package com.example.myapplicationexcel1

import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.util.Log
import com.github.mikephil.charting.charts.BarChart
import com.github.mikephil.charting.charts.CombinedChart
import com.github.mikephil.charting.charts.LineChart
import com.github.mikephil.charting.data.*
import com.github.mikephil.charting.utils.ColorTemplate
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.InputStream


class MainActivity : AppCompatActivity() {
    private lateinit var combinedChart: CombinedChart
    private lateinit var combinedData: CombinedData

    private lateinit var lineChartEntries: ArrayList<Entry>
    private lateinit var lineChartDataSet: LineDataSet
    private lateinit var lineData: LineData

    private lateinit var barChartEntries: ArrayList<BarEntry>
    private lateinit var barChartDataSet: BarDataSet
    private lateinit var barData: BarData

    private lateinit var data: InputStream
    private lateinit var wb: Workbook
    private lateinit var ws: Sheet

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        combinedChart = findViewById(R.id.combinedChart)
        initialize()

        getLineChartData()
        lineChartDataSet = LineDataSet(lineChartEntries, "First Line Chart")
        lineData = LineData(lineChartDataSet)


        getBarChartData()
        barChartDataSet = BarDataSet(barChartEntries, "Second Bar Chart")
        barChartDataSet.color = 1
        barData = BarData(barChartDataSet)


        combinedData = CombinedData()
        combinedData.setData(lineData)
        combinedData.setData(barData)
        combinedChart.data = combinedData
    }

    fun initialize() {
        Log.d("test", "Starting reading file")
        data = resources.openRawResource(R.raw.data)
        wb = WorkbookFactory.create(data)
    }

    fun getLineChartData() {
        ws = wb.getSheetAt(2)

        lineChartEntries = arrayListOf<Entry>()

        for (i in 1..ws.lastRowNum) {
            if (ws.getRow(i).getCell(0) == null) {
                break;
            }
            lineChartEntries.add(Entry(ws.getRow(i).getCell(0).numericCellValue.toFloat(),
                ws.getRow(i).getCell(1).numericCellValue.toFloat()
            ))
        }
    }

    fun getBarChartData() {
        // 4 - 8 columns
        ws = wb.getSheetAt(0)

        barChartEntries = arrayListOf<BarEntry>()

        val ebitdaSums = arrayListOf<Float>(0f, 0f, 0f, 0f, 0f)

        for (i in 1..ws.lastRowNum) {
            if (ws.getRow(i).getCell(0) == null) {
                break;
            }

            ebitdaSums[0] += if (ws.getRow(i).getCell(4) == null) 0f else ws.getRow(i).getCell(4).numericCellValue.toFloat()
            ebitdaSums[1] += if (ws.getRow(i).getCell(5) == null) 0f else ws.getRow(i).getCell(5).numericCellValue.toFloat()
            ebitdaSums[2] += if (ws.getRow(i).getCell(6) == null) 0f else ws.getRow(i).getCell(6).numericCellValue.toFloat()
            ebitdaSums[3] += if (ws.getRow(i).getCell(7) == null) 0f else ws.getRow(i).getCell(7).numericCellValue.toFloat()
            ebitdaSums[4] += if (ws.getRow(i).getCell(8) == null) 0f else ws.getRow(i).getCell(8).numericCellValue.toFloat()
        }

        for (i in 2018..2022) {
            barChartEntries.add(BarEntry(i.toFloat(), ebitdaSums[i - 2018]))
        }
    }
}