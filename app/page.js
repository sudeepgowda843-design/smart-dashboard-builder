"use client";

import { useMemo, useRef, useState } from "react";
import ExcelJS from "exceljs";
import Papa from "papaparse";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";

import {
  Upload,
  FileSpreadsheet,
  BarChart3,
  PieChart as PieChartIcon,
  LineChart as LineChartIcon,
  Table,
  AlertCircle,
  CheckCircle2,
  Download,
  Filter,
  RotateCcw,
  Sparkles,
  Database,
  Layers,
  Search,
  Plus,
  Trash2,
  FileDown,
  LayoutDashboard,
  Wand2,
} from "lucide-react";

import {
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
  XAxis,
  YAxis,
  Tooltip,
  CartesianGrid,
  ResponsiveContainer,
  Legend,
} from "recharts";

const COLORS = [
  "#2563eb",
  "#16a34a",
  "#dc2626",
  "#9333ea",
  "#ea580c",
  "#0891b2",
  "#4f46e5",
  "#be123c",
];

export default function Home() {
  const dashboardRef = useRef(null);

  const [fileName, setFileName] = useState("");
  const [rawData, setRawData] = useState([]);
  const [columns, setColumns] = useState([]);

  const [selectedDashboard, setSelectedDashboard] = useState("");
  const [selectedCategory, setSelectedCategory] = useState("");
  const [selectedNumber, setSelectedNumber] = useState("");
  const [selectedDate, setSelectedDate] = useState("");
  const [selectedStatus, setSelectedStatus] = useState("");

  const [filterColumn, setFilterColumn] = useState("");
  const [filterValue, setFilterValue] = useState("");
  const [searchText, setSearchText] = useState("");

  const [customChartTitle, setCustomChartTitle] = useState("");
  const [customChartType, setCustomChartType] = useState("bar");
  const [customXColumn, setCustomXColumn] = useState("");
  const [customYColumn, setCustomYColumn] = useState("");
  const [savedCharts, setSavedCharts] = useState([]);

  const [error, setError] = useState("");
  const [isExportingPdf, setIsExportingPdf] = useState(false);

  const handleFileUpload = async (event) => {
    setError("");

    const file = event.target.files[0];
    if (!file) return;

    setFileName(file.name);

    try {
      let cleanedData = [];
      const fileNameLower = file.name.toLowerCase();

      if (fileNameLower.endsWith(".xls")) {
        setError(
          "Old .xls files are not supported in Version 2.2. Please open the file in Excel and save it as .xlsx or CSV, then upload again."
        );
        resetData();
        return;
      }

      if (fileNameLower.endsWith(".csv")) {
        const text = await file.text();

        const parsed = Papa.parse(text, {
          header: true,
          skipEmptyLines: true,
        });

        cleanedData = parsed.data.map((row) => {
          const cleanedRow = {};

          Object.keys(row).forEach((key) => {
            const cleanKey = String(key).trim();
            cleanedRow[cleanKey] = row[key];
          });

          return cleanedRow;
        });
      } else if (fileNameLower.endsWith(".xlsx")) {
        const buffer = await file.arrayBuffer();

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.worksheets[0];

        if (!worksheet) {
          setError("No worksheet found in this Excel file.");
          resetData();
          return;
        }

        const headers = [];

        worksheet.getRow(1).eachCell((cell, colNumber) => {
          headers[colNumber] = String(cell.value || "").trim();
        });

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1) return;

          const rowData = {};

          headers.forEach((header, index) => {
            if (!header) return;

            const cell = row.getCell(index);
            let value = cell.value;

            if (value && typeof value === "object") {
              if (value instanceof Date) {
                value = value;
              } else if (value.text) {
                value = value.text;
              } else if (value.result) {
                value = value.result;
              } else if (value.richText) {
                value = value.richText.map((item) => item.text).join("");
              } else if (value.hyperlink) {
                value = value.text || value.hyperlink;
              } else {
                value = String(value);
              }
            }

            rowData[header] = value ?? "";
          });

          const hasAnyValue = Object.values(rowData).some(
            (value) => value !== "" && value !== null && value !== undefined
          );

          if (hasAnyValue) {
            cleanedData.push(rowData);
          }
        });
      } else {
        setError("Please upload only .xlsx or .csv files.");
        resetData();
        return;
      }

      if (!cleanedData || cleanedData.length === 0) {
        setError("The uploaded file is empty or could not be read.");
        resetData();
        return;
      }

      const detectedColumns = Object.keys(cleanedData[0]);

      setRawData(cleanedData);
      setColumns(detectedColumns);

      setSelectedDashboard("");
      setSelectedCategory("");
      setSelectedNumber("");
      setSelectedDate("");
      setSelectedStatus("");

      setFilterColumn("");
      setFilterValue("");
      setSearchText("");

      setCustomChartTitle("");
      setCustomChartType("bar");
      setCustomXColumn("");
      setCustomYColumn("");
      setSavedCharts([]);
    } catch (err) {
      console.error(err);
      setError("Could not read this file. Please upload a valid .xlsx or .csv file.");
      resetData();
    }
  };

  const resetData = () => {
    setRawData([]);
    setColumns([]);
    setSelectedDashboard("");
    setSelectedCategory("");
    setSelectedNumber("");
    setSelectedDate("");
    setSelectedStatus("");
    setFilterColumn("");
    setFilterValue("");
    setSearchText("");
    setCustomChartTitle("");
    setCustomChartType("bar");
    setCustomXColumn("");
    setCustomYColumn("");
    setSavedCharts([]);
  };

  const columnTypes = useMemo(() => {
    if (!rawData.length || !columns.length) {
      return {
        numberColumns: [],
        categoryColumns: [],
        dateColumns: [],
        statusColumns: [],
      };
    }

    const numberColumns = [];
    const categoryColumns = [];
    const dateColumns = [];
    const statusColumns = [];

    columns.forEach((column) => {
      const values = rawData
        .map((row) => row[column])
        .filter((value) => value !== "" && value !== null && value !== undefined);

      if (values.length === 0) {
        categoryColumns.push(column);
        return;
      }

      const numericCount = values.filter((value) => {
        if (typeof value === "number") return true;

        if (
          typeof value === "string" &&
          value.trim() !== "" &&
          !isNaN(Number(value))
        ) {
          return true;
        }

        return false;
      }).length;

      const dateCount = values.filter((value) => {
        if (value instanceof Date && !isNaN(value)) return true;

        if (typeof value === "string") {
          const parsed = Date.parse(value);
          return !isNaN(parsed) && value.match(/[0-9]/);
        }

        return false;
      }).length;

      const uniqueValues = [
        ...new Set(values.map((value) => String(value).toLowerCase().trim())),
      ];

      const statusKeywords = [
        "pending",
        "received",
        "paid",
        "unpaid",
        "present",
        "absent",
        "yes",
        "no",
        "completed",
        "not completed",
        "approved",
        "rejected",
        "active",
        "inactive",
      ];

      const hasStatusKeyword = uniqueValues.some((value) =>
        statusKeywords.some((keyword) => value.includes(keyword))
      );

      if (numericCount / values.length >= 0.8) {
        numberColumns.push(column);
      } else if (dateCount / values.length >= 0.7) {
        dateColumns.push(column);
      } else {
        categoryColumns.push(column);
      }

      if (hasStatusKeyword || uniqueValues.length <= 5) {
        statusColumns.push(column);
      }
    });

    return {
      numberColumns,
      categoryColumns,
      dateColumns,
      statusColumns,
    };
  }, [rawData, columns]);

  const filteredData = useMemo(() => {
    let data = [...rawData];

    if (filterColumn && filterValue) {
      data = data.filter(
        (row) => String(row[filterColumn] || "Missing") === filterValue
      );
    }

    if (searchText.trim()) {
      const search = searchText.toLowerCase().trim();

      data = data.filter((row) =>
        columns.some((column) =>
          String(row[column] || "").toLowerCase().includes(search)
        )
      );
    }

    return data;
  }, [rawData, filterColumn, filterValue, searchText, columns]);

  const filterValues = useMemo(() => {
    if (!filterColumn) return [];

    return [
      ...new Set(rawData.map((row) => String(row[filterColumn] || "Missing"))),
    ]
      .filter(Boolean)
      .sort();
  }, [rawData, filterColumn]);

  const dashboardSuggestions = useMemo(() => {
    const suggestions = [];

    if (columnTypes.categoryColumns.length && columnTypes.numberColumns.length) {
      suggestions.push({
        id: "category-performance",
        title: "Category-wise Performance Dashboard",
        description:
          "Best for sales, scholarship amount, placement count, expenses, or department-wise analysis.",
        score: 95,
        required: "Category + Numeric",
      });
    }

    if (columnTypes.statusColumns.length) {
      suggestions.push({
        id: "status-summary",
        title: "Status Summary Dashboard",
        description:
          "Best for pending vs completed, paid vs unpaid, present vs absent, approved vs rejected.",
        score: 90,
        required: "Status column",
      });
    }

    if (columnTypes.dateColumns.length && columnTypes.numberColumns.length) {
      suggestions.push({
        id: "time-trend",
        title: "Time Trend Dashboard",
        description: "Best for tracking change over time using date and numeric values.",
        score: 85,
        required: "Date + Numeric",
      });
    }

    if (columnTypes.categoryColumns.length) {
      suggestions.push({
        id: "category-count",
        title: "Category Count Dashboard",
        description:
          "Best for department count, section count, region count, event participation, or category split.",
        score: 80,
        required: "Category column",
      });
    }

    suggestions.push({
      id: "full-overview",
      title: "Full Data Overview Dashboard",
      description: "Best for checking missing values, duplicates, and overall data health.",
      score: 100,
      required: "Any data file",
    });

    return suggestions;
  }, [columnTypes]);

  const missingValuesCount = useMemo(() => {
    if (!filteredData.length) return 0;

    let count = 0;

    filteredData.forEach((row) => {
      columns.forEach((column) => {
        if (row[column] === "" || row[column] === null || row[column] === undefined) {
          count += 1;
        }
      });
    });

    return count;
  }, [filteredData, columns]);

  const duplicateRowsCount = useMemo(() => {
    if (!filteredData.length) return 0;

    const seen = new Set();
    let duplicates = 0;

    filteredData.forEach((row) => {
      const key = JSON.stringify(row);

      if (seen.has(key)) {
        duplicates += 1;
      } else {
        seen.add(key);
      }
    });

    return duplicates;
  }, [filteredData]);

  const dataQualityScore = useMemo(() => {
    if (!filteredData.length || !columns.length) return 0;

    const totalCells = filteredData.length * columns.length;
    if (totalCells === 0) return 0;

    const missingPercentage = (missingValuesCount / totalCells) * 100;
    const duplicatePercentage = (duplicateRowsCount / filteredData.length) * 100;

    const score = Math.max(0, 100 - missingPercentage - duplicatePercentage);

    return Math.round(score);
  }, [filteredData, columns, missingValuesCount, duplicateRowsCount]);

  const ruleBasedInsights = useMemo(() => {
    const insights = [];

    if (!filteredData.length) return insights;

    insights.push(
      `Current view contains ${filteredData.length} records and ${columns.length} columns.`
    );

    if (missingValuesCount > 0) {
      insights.push(
        `There are ${missingValuesCount} missing values in the current filtered data.`
      );
    } else {
      insights.push("No missing values found in the current filtered data.");
    }

    if (duplicateRowsCount > 0) {
      insights.push(`${duplicateRowsCount} duplicate rows were found.`);
    } else {
      insights.push("No duplicate rows found in the current filtered data.");
    }

    if (dataQualityScore >= 90) {
      insights.push("The data quality score is strong and suitable for reporting.");
    } else if (dataQualityScore >= 70) {
      insights.push(
        "The data quality score is acceptable, but cleaning can improve reporting accuracy."
      );
    } else {
      insights.push(
        "The data quality score is low. Review missing values and duplicate rows before reporting."
      );
    }

    if (columnTypes.categoryColumns[0]) {
      const category = columnTypes.categoryColumns[0];
      const grouped = getGroupedCountDataHelper(category, filteredData);

      if (grouped[0]) {
        insights.push(
          `Most frequent ${category}: ${grouped[0].name} with ${grouped[0].value} records.`
        );
      }
    }

    return insights;
  }, [
    filteredData,
    columns,
    missingValuesCount,
    duplicateRowsCount,
    dataQualityScore,
    columnTypes.categoryColumns,
  ]);

  const applyDashboardDefaults = (dashboard) => {
    setSelectedDashboard(dashboard.title);

    if (dashboard.id === "category-performance") {
      setSelectedCategory(columnTypes.categoryColumns[0] || "");
      setSelectedNumber(columnTypes.numberColumns[0] || "");
    }

    if (dashboard.id === "status-summary") {
      setSelectedStatus(columnTypes.statusColumns[0] || "");
    }

    if (dashboard.id === "time-trend") {
      setSelectedDate(columnTypes.dateColumns[0] || "");
      setSelectedNumber(columnTypes.numberColumns[0] || "");
    }

    if (dashboard.id === "category-count") {
      setSelectedCategory(columnTypes.categoryColumns[0] || "");
    }
  };

  const getGroupedSumData = (categoryColumn, numberColumn, data = filteredData) => {
    return getGroupedSumDataHelper(categoryColumn, numberColumn, data);
  };

  const getGroupedCountData = (categoryColumn, data = filteredData) => {
    return getGroupedCountDataHelper(categoryColumn, data);
  };

  const getTimeTrendData = (dateColumn, numberColumn, data = filteredData) => {
    return getTimeTrendDataHelper(dateColumn, numberColumn, data);
  };

  const getCategoryNumberData = () => {
    return getGroupedSumData(selectedCategory, selectedNumber);
  };

  const getCategoryCountData = () => {
    return getGroupedCountData(selectedCategory);
  };

  const getStatusData = () => {
    return getGroupedCountData(selectedStatus);
  };

  const getPresetTimeTrendData = () => {
    return getTimeTrendData(selectedDate, selectedNumber);
  };

  const addCustomChart = () => {
    if (!customXColumn) {
      setError("Please select an X-axis/category column before adding a chart.");
      return;
    }

    if (
      ["bar", "horizontal-bar", "line"].includes(customChartType) &&
      !customYColumn
    ) {
      setError("Please select a numeric Y-axis column for this chart type.");
      return;
    }

    setError("");

    const chart = {
      id: crypto.randomUUID(),
      title:
        customChartTitle ||
        `${getChartTypeLabel(customChartType)}: ${customXColumn}${
          customYColumn ? ` by ${customYColumn}` : ""
        }`,
      type: customChartType,
      xColumn: customXColumn,
      yColumn: customYColumn,
    };

    setSavedCharts((prev) => [...prev, chart]);
    setCustomChartTitle("");
  };

  const removeCustomChart = (chartId) => {
    setSavedCharts((prev) => prev.filter((chart) => chart.id !== chartId));
  };

  const clearCustomDashboard = () => {
    setSavedCharts([]);
  };

  const downloadCSV = () => {
    if (!filteredData.length) return;

    const csv = Papa.unparse(filteredData);

    const blob = new Blob([csv], {
      type: "text/csv;charset=utf-8;",
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = "dashpilot_filtered_data.csv";
    link.click();

    URL.revokeObjectURL(url);
  };

  const downloadExcel = async () => {
    if (!filteredData.length) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Filtered Data");

    worksheet.columns = columns.map((column) => ({
      header: column,
      key: column,
      width: 22,
    }));

    filteredData.forEach((row) => {
      worksheet.addRow(row);
    });

    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "D9EAF7" },
    };

    worksheet.views = [{ state: "frozen", ySplit: 1 }];

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");

    link.href = url;
    link.download = "dashpilot_filtered_data.xlsx";
    link.click();

    URL.revokeObjectURL(url);
  };

  const downloadDashboardPDF = async () => {
    if (!dashboardRef.current) return;

    try {
      setIsExportingPdf(true);
      setError("");

      const sourceNode = dashboardRef.current;
      const clonedNode = sourceNode.cloneNode(true);

      const exportWrapper = document.createElement("div");

      exportWrapper.style.position = "absolute";
      exportWrapper.style.left = "-99999px";
      exportWrapper.style.top = "0";
      exportWrapper.style.width = "1200px";
      exportWrapper.style.background = "#f8fafc";
      exportWrapper.style.padding = "24px";
      exportWrapper.style.fontFamily = "Arial, Helvetica, sans-serif";
      exportWrapper.style.color = "#0f172a";

      clonedNode.style.width = "100%";
      clonedNode.style.background = "#f8fafc";
      clonedNode.style.color = "#0f172a";

      const allElements = clonedNode.querySelectorAll("*");

      allElements.forEach((element) => {
        const tagName = element.tagName.toLowerCase();

        element.removeAttribute("class");

        const isSvgElement = [
          "svg",
          "path",
          "circle",
          "rect",
          "line",
          "polyline",
          "polygon",
          "text",
          "g",
        ].includes(tagName);

        if (isSvgElement) {
          const fill = element.getAttribute("fill");
          const stroke = element.getAttribute("stroke");

          if (
            fill &&
            (fill.includes("lab") ||
              fill.includes("lch") ||
              fill.includes("oklab") ||
              fill.includes("oklch") ||
              fill.includes("color("))
          ) {
            element.setAttribute("fill", "#2563eb");
          }

          if (
            stroke &&
            (stroke.includes("lab") ||
              stroke.includes("lch") ||
              stroke.includes("oklab") ||
              stroke.includes("oklch") ||
              stroke.includes("color("))
          ) {
            element.setAttribute("stroke", "#2563eb");
          }

          return;
        }

        element.removeAttribute("style");

        element.style.boxSizing = "border-box";
        element.style.color = "#0f172a";
        element.style.borderColor = "#e2e8f0";
        element.style.outlineColor = "#e2e8f0";
        element.style.textDecorationColor = "#0f172a";
        element.style.boxShadow = "none";
        element.style.textShadow = "none";
        element.style.backgroundImage = "none";

        if (tagName === "div" || tagName === "section" || tagName === "main") {
          element.style.backgroundColor = "#ffffff";
          element.style.border = "1px solid #e2e8f0";
          element.style.borderRadius = "18px";
          element.style.padding = "16px";
          element.style.marginBottom = "16px";
        }

        if (tagName === "h1") {
          element.style.fontSize = "34px";
          element.style.fontWeight = "900";
          element.style.margin = "0 0 12px 0";
          element.style.color = "#0f172a";
        }

        if (tagName === "h2") {
          element.style.fontSize = "26px";
          element.style.fontWeight = "900";
          element.style.margin = "0 0 10px 0";
          element.style.color = "#0f172a";
        }

        if (tagName === "h3") {
          element.style.fontSize = "20px";
          element.style.fontWeight = "900";
          element.style.margin = "0 0 10px 0";
          element.style.color = "#0f172a";
        }

        if (tagName === "h4") {
          element.style.fontSize = "16px";
          element.style.fontWeight = "800";
          element.style.margin = "0 0 8px 0";
          element.style.color = "#334155";
        }

        if (tagName === "p" || tagName === "span" || tagName === "label") {
          element.style.fontSize = "13px";
          element.style.lineHeight = "1.5";
          element.style.color = "#475569";
        }

        if (tagName === "button") {
          element.style.display = "none";
        }

        if (tagName === "input" || tagName === "select") {
          element.style.display = "none";
        }

        if (tagName === "table") {
          element.style.width = "100%";
          element.style.borderCollapse = "collapse";
          element.style.backgroundColor = "#ffffff";
        }

        if (tagName === "thead") {
          element.style.backgroundColor = "#0f172a";
        }

        if (tagName === "th") {
          element.style.backgroundColor = "#0f172a";
          element.style.color = "#ffffff";
          element.style.padding = "10px";
          element.style.border = "1px solid #e2e8f0";
          element.style.fontSize = "11px";
          element.style.textAlign = "left";
        }

        if (tagName === "td") {
          element.style.padding = "9px";
          element.style.border = "1px solid #e2e8f0";
          element.style.fontSize = "11px";
          element.style.color = "#334155";
        }
      });

      exportWrapper.appendChild(clonedNode);
      document.body.appendChild(exportWrapper);

      const canvas = await html2canvas(exportWrapper, {
        scale: 2,
        backgroundColor: "#f8fafc",
        useCORS: true,
        logging: false,
      });

      document.body.removeChild(exportWrapper);

      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF("p", "mm", "a4");

      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();

      const imgWidth = pageWidth;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;

      let heightLeft = imgHeight;
      let position = 0;

      pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
      heightLeft -= pageHeight;

      while (heightLeft > 0) {
        position = heightLeft - imgHeight;
        pdf.addPage();
        pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;
      }

      pdf.save("dashpilot_dashboard.pdf");
    } catch (err) {
      console.error(err);
      setError(
        "Could not export PDF. The app is working, but the PDF exporter could not process some browser-rendered styles."
      );
    } finally {
      setIsExportingPdf(false);
    }
  };

  const clearFilters = () => {
    setFilterColumn("");
    setFilterValue("");
    setSearchText("");
  };

  const renderDashboard = () => {
    if (!selectedDashboard) {
      return (
        <div className="rounded-3xl border border-dashed border-slate-300 bg-white/90 p-10 text-center shadow-xl shadow-slate-200/70 backdrop-blur">
          <Sparkles className="mx-auto mb-3 h-10 w-10 text-blue-500" />
          <h3 className="text-xl font-black text-slate-950">
            Select a dashboard card
          </h3>
          <p className="mx-auto mt-2 max-w-xl text-sm text-slate-500">
            Version 2.2 recommends dashboards, builds charts, and supports PDF export.
          </p>
        </div>
      );
    }

    if (selectedDashboard === "Category-wise Performance Dashboard") {
      const chartData = getCategoryNumberData();

      return (
        <DashboardCard title="Category-wise Performance Dashboard">
          <SelectorGrid>
            <SelectBox
              label="Category column"
              value={selectedCategory}
              onChange={setSelectedCategory}
              options={columnTypes.categoryColumns}
            />

            <SelectBox
              label="Numeric column"
              value={selectedNumber}
              onChange={setSelectedNumber}
              options={columnTypes.numberColumns}
            />
          </SelectorGrid>

          {chartData.length > 0 && (
            <>
              <InsightBox
                text={`Top category: ${chartData[0]?.name} with value ${chartData[0]?.value}.`}
              />

              <div className="mt-6 grid gap-6 xl:grid-cols-2">
                <ChartPanel title={`${selectedNumber} by ${selectedCategory}`}>
                  <VerticalBar data={chartData} color="#2563eb" />
                </ChartPanel>

                <ChartPanel title="Distribution">
                  <PieVisual data={chartData} />
                </ChartPanel>
              </div>
            </>
          )}
        </DashboardCard>
      );
    }

    if (selectedDashboard === "Status Summary Dashboard") {
      const chartData = getStatusData();

      return (
        <DashboardCard title="Status Summary Dashboard">
          <SelectorGrid>
            <SelectBox
              label="Status column"
              value={selectedStatus}
              onChange={setSelectedStatus}
              options={columnTypes.statusColumns}
            />
          </SelectorGrid>

          {chartData.length > 0 && (
            <>
              <InsightBox
                text={`Most common status: ${chartData[0]?.name} with ${chartData[0]?.value} records.`}
              />

              <div className="mt-6 grid gap-6 xl:grid-cols-2">
                <ChartPanel title={`${selectedStatus} Count`}>
                  <VerticalBar data={chartData} color="#16a34a" />
                </ChartPanel>

                <ChartPanel title={`${selectedStatus} Distribution`}>
                  <PieVisual data={chartData} />
                </ChartPanel>
              </div>
            </>
          )}
        </DashboardCard>
      );
    }

    if (selectedDashboard === "Time Trend Dashboard") {
      const chartData = getPresetTimeTrendData();

      return (
        <DashboardCard title="Time Trend Dashboard">
          <SelectorGrid>
            <SelectBox
              label="Date column"
              value={selectedDate}
              onChange={setSelectedDate}
              options={columnTypes.dateColumns}
            />

            <SelectBox
              label="Numeric column"
              value={selectedNumber}
              onChange={setSelectedNumber}
              options={columnTypes.numberColumns}
            />
          </SelectorGrid>

          {chartData.length > 0 && (
            <>
              <InsightBox
                text={`This chart shows ${selectedNumber} movement over ${chartData.length} date points.`}
              />

              <div className="mt-6">
                <ChartPanel title={`${selectedNumber} Trend Over Time`}>
                  <LineVisual data={chartData} />
                </ChartPanel>
              </div>
            </>
          )}
        </DashboardCard>
      );
    }

    if (selectedDashboard === "Category Count Dashboard") {
      const chartData = getCategoryCountData();

      return (
        <DashboardCard title="Category Count Dashboard">
          <SelectorGrid>
            <SelectBox
              label="Category column"
              value={selectedCategory}
              onChange={setSelectedCategory}
              options={columnTypes.categoryColumns}
            />
          </SelectorGrid>

          {chartData.length > 0 && (
            <>
              <InsightBox
                text={`Highest count: ${chartData[0]?.name} with ${chartData[0]?.value} records.`}
              />

              <div className="mt-6 grid gap-6 xl:grid-cols-2">
                <ChartPanel title={`Count by ${selectedCategory}`}>
                  <VerticalBar data={chartData} color="#9333ea" />
                </ChartPanel>

                <ChartPanel title={`${selectedCategory} Distribution`}>
                  <PieVisual data={chartData} />
                </ChartPanel>
              </div>
            </>
          )}
        </DashboardCard>
      );
    }

    if (selectedDashboard === "Full Data Overview Dashboard") {
      const missingData = columns
        .map((column) => {
          const count = filteredData.filter(
            (row) =>
              row[column] === "" ||
              row[column] === null ||
              row[column] === undefined
          ).length;

          return {
            name: column,
            value: count,
          };
        })
        .filter((item) => item.value > 0);

      return (
        <DashboardCard title="Full Data Overview Dashboard">
          <InsightBox
            text={`Data quality score is ${dataQualityScore}/100 based on missing values and duplicate rows in the current filtered data.`}
          />

          {missingData.length > 0 ? (
            <div className="mt-6">
              <ChartPanel title="Missing Values by Column">
                <VerticalBar data={missingData} color="#dc2626" />
              </ChartPanel>
            </div>
          ) : (
            <div className="mt-6 rounded-3xl bg-green-50 p-4 font-semibold text-green-700">
              No missing values found in the current filtered data.
            </div>
          )}
        </DashboardCard>
      );
    }

    return null;
  };

  return (
    <main className="min-h-screen">
      <section className="relative overflow-hidden border-b border-white/60 bg-white/80 backdrop-blur-xl">
        <div className="absolute inset-0 -z-10 bg-[radial-gradient(circle_at_top_left,_rgba(37,99,235,0.18),_transparent_35%),radial-gradient(circle_at_top_right,_rgba(147,51,234,0.14),_transparent_35%)]" />

        <div className="mx-auto grid max-w-7xl gap-8 px-6 py-10 lg:grid-cols-[1.2fr_0.8fr] lg:items-center">
          <div>
            <div className="mb-4 inline-flex items-center gap-2 rounded-full border border-blue-100 bg-white px-4 py-2 text-sm font-bold text-blue-700 shadow-sm">
              <Sparkles className="h-4 w-4" />
              DashPilot Version 2.2
            </div>

            <h1 className="max-w-3xl text-4xl font-black tracking-tight text-slate-950 md:text-6xl">
              Turn spreadsheets into beautiful dashboards.
            </h1>

            <p className="mt-5 max-w-2xl text-base leading-7 text-slate-600 md:text-lg">
              Upload Excel or CSV files, detect columns automatically, generate
              dashboards, build custom chart layouts, and export reports.
            </p>

            <div className="mt-6 flex flex-wrap gap-3 text-sm font-semibold text-slate-600">
              <span className="rounded-full bg-white px-4 py-2 shadow-sm">
                Custom builder
              </span>
              <span className="rounded-full bg-white px-4 py-2 shadow-sm">
                PDF export
              </span>
              <span className="rounded-full bg-white px-4 py-2 shadow-sm">
                No API key
              </span>
              <span className="rounded-full bg-white px-4 py-2 shadow-sm">
                Vercel ready
              </span>
            </div>
          </div>

          <div className="rounded-[2rem] border border-white bg-white/90 p-6 shadow-2xl shadow-blue-100">
            <div className="mb-5 flex h-14 w-14 items-center justify-center rounded-2xl bg-blue-600 text-white shadow-lg shadow-blue-200">
              <Upload className="h-7 w-7" />
            </div>

            <h2 className="text-xl font-black text-slate-950">
              Upload your data file
            </h2>

            <p className="mt-2 text-sm leading-6 text-slate-500">
              Supported formats are .xlsx and .csv. Old .xls files must be saved as
              .xlsx before uploading.
            </p>

            <label className="mt-6 flex cursor-pointer flex-col items-center justify-center rounded-3xl border-2 border-dashed border-blue-200 bg-blue-50/70 px-6 py-10 text-center transition hover:border-blue-500 hover:bg-blue-50">
              <div className="mb-3 rounded-2xl bg-white p-4 shadow-sm">
                <FileSpreadsheet className="h-8 w-8 text-blue-600" />
              </div>

              <span className="text-base font-black text-slate-900">
                Click to upload Excel/CSV
              </span>

              <span className="mt-1 text-sm text-slate-500">
                Your file is processed in the browser
              </span>

              <input
                type="file"
                accept=".xlsx,.csv"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>
          </div>
        </div>
      </section>

      <section className="mx-auto max-w-7xl px-6 py-8">
        {error && (
          <div className="mb-6 flex items-center gap-3 rounded-2xl border border-red-200 bg-red-50 p-4 text-red-700 shadow-sm">
            <AlertCircle className="h-5 w-5" />
            {error}
          </div>
        )}

        {!rawData.length ? (
          <EmptyState />
        ) : (
          <>
            <div className="mb-6 rounded-[2rem] border border-white bg-white/90 p-5 shadow-xl shadow-slate-200/70 backdrop-blur">
              <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
                <div className="flex items-center gap-4">
                  <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-gradient-to-br from-blue-600 to-indigo-600 text-white shadow-lg shadow-blue-200">
                    <FileSpreadsheet className="h-7 w-7" />
                  </div>

                  <div>
                    <h2 className="text-lg font-black text-slate-950">
                      {fileName}
                    </h2>
                    <p className="text-sm text-slate-500">
                      Showing {filteredData.length} of {rawData.length} records
                    </p>
                  </div>
                </div>

                <div className="flex flex-wrap gap-2">
                  <button
                    onClick={downloadCSV}
                    className="flex items-center gap-2 rounded-2xl bg-slate-950 px-5 py-3 text-sm font-bold text-white shadow-lg shadow-slate-200 transition hover:-translate-y-0.5 hover:bg-slate-800"
                  >
                    <Download className="h-4 w-4" />
                    Export CSV
                  </button>

                  <button
                    onClick={downloadExcel}
                    className="flex items-center gap-2 rounded-2xl bg-green-600 px-5 py-3 text-sm font-bold text-white shadow-lg shadow-green-100 transition hover:-translate-y-0.5 hover:bg-green-700"
                  >
                    <Download className="h-4 w-4" />
                    Export Excel
                  </button>

                  <button
                    onClick={downloadDashboardPDF}
                    disabled={isExportingPdf}
                    className="flex items-center gap-2 rounded-2xl bg-blue-600 px-5 py-3 text-sm font-bold text-white shadow-lg shadow-blue-100 transition hover:-translate-y-0.5 hover:bg-blue-700 disabled:cursor-not-allowed disabled:opacity-60"
                  >
                    <FileDown className="h-4 w-4" />
                    {isExportingPdf ? "Exporting..." : "Export PDF"}
                  </button>
                </div>
              </div>
            </div>

            <div ref={dashboardRef}>
              <div className="mb-8 grid gap-4 md:grid-cols-2 lg:grid-cols-5">
                <MetricCard title="Rows Shown" value={filteredData.length} />
                <MetricCard title="Total Columns" value={columns.length} />
                <MetricCard title="Missing Values" value={missingValuesCount} />
                <MetricCard title="Duplicate Rows" value={duplicateRowsCount} />
                <MetricCard title="Quality Score" value={`${dataQualityScore}/100`} />
              </div>

              <InsightsPanel insights={ruleBasedInsights} />

              <FilterPanel
                columns={columns}
                filterColumn={filterColumn}
                setFilterColumn={setFilterColumn}
                filterValue={filterValue}
                setFilterValue={setFilterValue}
                filterValues={filterValues}
                searchText={searchText}
                setSearchText={setSearchText}
                clearFilters={clearFilters}
              />

              <div className="mt-8 grid gap-8 lg:grid-cols-[390px_1fr]">
                <aside className="space-y-6 lg:sticky lg:top-6 lg:self-start">
                  <DetectedColumnsCard columnTypes={columnTypes} />

                  <div className="rounded-[2rem] border border-white bg-white/90 p-5 shadow-xl shadow-slate-200/70 backdrop-blur">
                    <h3 className="mb-4 flex items-center gap-2 font-black text-slate-950">
                      <Layers className="h-5 w-5 text-blue-600" />
                      Recommended Dashboards
                    </h3>

                    <div className="space-y-3">
                      {dashboardSuggestions.map((dashboard) => (
                        <DashboardSuggestionCard
                          key={dashboard.id}
                          dashboard={dashboard}
                          active={selectedDashboard === dashboard.title}
                          onClick={() => applyDashboardDefaults(dashboard)}
                        />
                      ))}
                    </div>
                  </div>

                  <CustomDashboardBuilder
                    customChartTitle={customChartTitle}
                    setCustomChartTitle={setCustomChartTitle}
                    customChartType={customChartType}
                    setCustomChartType={setCustomChartType}
                    customXColumn={customXColumn}
                    setCustomXColumn={setCustomXColumn}
                    customYColumn={customYColumn}
                    setCustomYColumn={setCustomYColumn}
                    columns={columns}
                    numberColumns={columnTypes.numberColumns}
                    addCustomChart={addCustomChart}
                    clearCustomDashboard={clearCustomDashboard}
                    savedCharts={savedCharts}
                  />
                </aside>

                <div className="space-y-8">
                  {renderDashboard()}

                  <CustomDashboardWorkspace
                    savedCharts={savedCharts}
                    removeCustomChart={removeCustomChart}
                    getChartDataForSavedChart={(chart) =>
                      getChartDataForSavedChart(chart, filteredData)
                    }
                  />

                  <DataPreview columns={columns} data={filteredData} />
                </div>
              </div>
            </div>
          </>
        )}
      </section>
    </main>
  );
}

function EmptyState() {
  return (
    <div className="rounded-[2rem] border border-white bg-white/90 p-10 text-center shadow-xl shadow-slate-200/70 backdrop-blur">
      <div className="mx-auto mb-5 flex h-20 w-20 items-center justify-center rounded-3xl bg-gradient-to-br from-blue-600 to-indigo-600 text-white shadow-xl shadow-blue-200">
        <LayoutDashboard className="h-10 w-10" />
      </div>

      <h2 className="text-3xl font-black text-slate-950">
        Your dashboard workspace is ready
      </h2>

      <p className="mx-auto mt-3 max-w-2xl text-base leading-7 text-slate-600">
        Upload an Excel or CSV file to generate dashboard recommendations, KPI
        cards, filters, custom charts, PDF reports, and export-ready data.
      </p>

      <div className="mx-auto mt-8 grid max-w-4xl gap-4 md:grid-cols-3">
        <FeatureCard
          icon={<BarChart3 className="h-6 w-6" />}
          title="Custom Dashboard"
          description="Build multiple charts in one dashboard."
        />
        <FeatureCard
          icon={<Wand2 className="h-6 w-6" />}
          title="Smart Insights"
          description="Get rule-based insights without API cost."
        />
        <FeatureCard
          icon={<FileDown className="h-6 w-6" />}
          title="PDF Export"
          description="Export your dashboard as a PDF report."
        />
      </div>
    </div>
  );
}

function MetricCard({ title, value }) {
  return (
    <div className="group rounded-[1.75rem] border border-white bg-white/90 p-5 shadow-xl shadow-slate-200/60 transition hover:-translate-y-1 hover:shadow-2xl">
      <p className="text-sm font-bold text-slate-500">{title}</p>
      <h3 className="mt-3 text-3xl font-black tracking-tight text-slate-950">
        {value}
      </h3>
      <div className="mt-4 h-1.5 rounded-full bg-slate-100">
        <div className="h-1.5 w-2/3 rounded-full bg-gradient-to-r from-blue-600 to-indigo-600 transition group-hover:w-full" />
      </div>
    </div>
  );
}

function InsightsPanel({ insights }) {
  return (
    <div className="mb-8 rounded-[2rem] border border-white bg-white/90 p-5 shadow-xl shadow-slate-200/70 backdrop-blur">
      <div className="mb-4 flex items-center gap-2">
        <Wand2 className="h-5 w-5 text-blue-600" />
        <h3 className="font-black text-slate-950">Smart Insights</h3>
      </div>

      <div className="grid gap-3 md:grid-cols-2">
        {insights.map((insight, index) => (
          <div
            key={index}
            className="rounded-2xl border border-blue-100 bg-blue-50/70 p-4 text-sm font-bold leading-6 text-blue-900"
          >
            {insight}
          </div>
        ))}
      </div>
    </div>
  );
}

function FilterPanel({
  columns,
  filterColumn,
  setFilterColumn,
  filterValue,
  setFilterValue,
  filterValues,
  searchText,
  setSearchText,
  clearFilters,
}) {
  return (
    <div className="rounded-[2rem] border border-white bg-white/90 p-5 shadow-xl shadow-slate-200/70 backdrop-blur">
      <div className="mb-5 flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
        <div>
          <h3 className="flex items-center gap-2 font-black text-slate-950">
            <Filter className="h-5 w-5 text-blue-600" />
            Global Filters
          </h3>
          <p className="mt-1 text-sm text-slate-500">
            Filter the entire dashboard before creating charts.
          </p>
        </div>

        <button
          onClick={clearFilters}
          className="flex items-center justify-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-bold text-slate-700 transition hover:bg-slate-50"
        >
          <RotateCcw className="h-4 w-4" />
          Reset Filters
        </button>
      </div>

      <div className="grid gap-4 lg:grid-cols-3">
        <div>
          <label className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">
            Filter column
          </label>
          <select
            value={filterColumn}
            onChange={(event) => {
              setFilterColumn(event.target.value);
              setFilterValue("");
            }}
            className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-semibold outline-none transition focus:border-blue-500 focus:bg-white focus:ring-4 focus:ring-blue-100"
          >
            <option value="">No filter</option>
            {columns.map((column) => (
              <option key={column} value={column}>
                {column}
              </option>
            ))}
          </select>
        </div>

        <div>
          <label className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">
            Filter value
          </label>
          <select
            value={filterValue}
            onChange={(event) => setFilterValue(event.target.value)}
            disabled={!filterColumn}
            className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-semibold outline-none transition focus:border-blue-500 focus:bg-white focus:ring-4 focus:ring-blue-100 disabled:cursor-not-allowed disabled:opacity-60"
          >
            <option value="">All values</option>
            {filterValues.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        </div>

        <div>
          <label className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">
            Search all data
          </label>
          <div className="flex items-center rounded-2xl border border-slate-200 bg-slate-50 px-4 transition focus-within:border-blue-500 focus-within:bg-white focus-within:ring-4 focus-within:ring-blue-100">
            <Search className="h-4 w-4 text-slate-400" />
            <input
              value={searchText}
              onChange={(event) => setSearchText(event.target.value)}
              placeholder="Search anything..."
              className="w-full bg-transparent px-3 py-3 text-sm font-semibold outline-none"
            />
          </div>
        </div>
      </div>
    </div>
  );
}

function DetectedColumnsCard({ columnTypes }) {
  return (
    <div className="rounded-[2rem] border border-white bg-white/90 p-5 shadow-xl shadow-slate-200/70 backdrop-blur">
      <h3 className="mb-4 flex items-center gap-2 font-black text-slate-950">
        <Database className="h-5 w-5 text-blue-600" />
        Detected Columns
      </h3>

      <ColumnGroup title="Numeric Columns" items={columnTypes.numberColumns} />
      <ColumnGroup title="Date Columns" items={columnTypes.dateColumns} />
      <ColumnGroup title="Category Columns" items={columnTypes.categoryColumns} />
      <ColumnGroup title="Status Columns" items={columnTypes.statusColumns} />
    </div>
  );
}

function DashboardSuggestionCard({ dashboard, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className={`group w-full rounded-3xl border p-4 text-left transition hover:-translate-y-0.5 ${
        active
          ? "border-blue-500 bg-gradient-to-br from-blue-50 to-indigo-50 shadow-lg shadow-blue-100"
          : "border-slate-200 bg-white hover:border-blue-300 hover:shadow-lg hover:shadow-slate-200"
      }`}
    >
      <div className="mb-3 flex items-start justify-between gap-3">
        <div className="flex items-center gap-3">
          <div
            className={`flex h-10 w-10 items-center justify-center rounded-2xl ${
              active ? "bg-blue-600 text-white" : "bg-slate-100"
            }`}
          >
            {getDashboardIcon(dashboard.title)}
          </div>

          <h4 className="text-sm font-black leading-5 text-slate-950">
            {dashboard.title}
          </h4>
        </div>

        <span className="rounded-full bg-green-50 px-2.5 py-1 text-xs font-black text-green-700">
          {dashboard.score}%
        </span>
      </div>

      <p className="text-xs leading-5 text-slate-500">{dashboard.description}</p>

      <div className="mt-3 flex items-center justify-between rounded-2xl bg-slate-50 px-3 py-2 text-xs font-bold text-slate-600">
        <span>Needs: {dashboard.required}</span>
        <span className="text-blue-600">Create →</span>
      </div>
    </button>
  );
}

function CustomDashboardBuilder({
  customChartTitle,
  setCustomChartTitle,
  customChartType,
  setCustomChartType,
  customXColumn,
  setCustomXColumn,
  customYColumn,
  setCustomYColumn,
  columns,
  numberColumns,
  addCustomChart,
  clearCustomDashboard,
  savedCharts,
}) {
  const needsYColumn = ["bar", "horizontal-bar", "line"].includes(customChartType);

  return (
    <div className="rounded-[2rem] border border-white bg-white/90 p-5 shadow-xl shadow-slate-200/70 backdrop-blur">
      <h3 className="mb-2 flex items-center gap-2 font-black text-slate-950">
        <LayoutDashboard className="h-5 w-5 text-blue-600" />
        Custom Builder
      </h3>

      <p className="mb-5 text-sm leading-6 text-slate-500">
        Add multiple charts and build your own dashboard layout.
      </p>

      <div className="space-y-4">
        <div>
          <label className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">
            Chart title
          </label>
          <input
            value={customChartTitle}
            onChange={(event) => setCustomChartTitle(event.target.value)}
            placeholder="Optional title"
            className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-semibold outline-none transition focus:border-blue-500 focus:bg-white focus:ring-4 focus:ring-blue-100"
          />
        </div>

        <div>
          <label className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">
            Chart type
          </label>
          <select
            value={customChartType}
            onChange={(event) => setCustomChartType(event.target.value)}
            className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-semibold outline-none transition focus:border-blue-500 focus:bg-white focus:ring-4 focus:ring-blue-100"
          >
            <option value="bar">Bar chart</option>
            <option value="horizontal-bar">Horizontal bar chart</option>
            <option value="pie">Pie chart</option>
            <option value="line">Line chart</option>
          </select>
        </div>

        <div>
          <label className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">
            X-axis / category
          </label>
          <select
            value={customXColumn}
            onChange={(event) => setCustomXColumn(event.target.value)}
            className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-semibold outline-none transition focus:border-blue-500 focus:bg-white focus:ring-4 focus:ring-blue-100"
          >
            <option value="">Select column</option>
            {columns.map((column) => (
              <option key={column} value={column}>
                {column}
              </option>
            ))}
          </select>
        </div>

        {needsYColumn && (
          <div>
            <label className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">
              Y-axis / numeric value
            </label>
            <select
              value={customYColumn}
              onChange={(event) => setCustomYColumn(event.target.value)}
              className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-semibold outline-none transition focus:border-blue-500 focus:bg-white focus:ring-4 focus:ring-blue-100"
            >
              <option value="">Select numeric column</option>
              {numberColumns.map((column) => (
                <option key={column} value={column}>
                  {column}
                </option>
              ))}
            </select>
          </div>
        )}

        <button
          onClick={addCustomChart}
          className="flex w-full items-center justify-center gap-2 rounded-2xl bg-blue-600 px-4 py-3 text-sm font-black text-white shadow-lg shadow-blue-100 transition hover:bg-blue-700"
        >
          <Plus className="h-4 w-4" />
          Add Chart
        </button>

        {savedCharts.length > 0 && (
          <button
            onClick={clearCustomDashboard}
            className="flex w-full items-center justify-center gap-2 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm font-black text-red-700 transition hover:bg-red-100"
          >
            <Trash2 className="h-4 w-4" />
            Clear Custom Dashboard
          </button>
        )}
      </div>
    </div>
  );
}

function CustomDashboardWorkspace({
  savedCharts,
  removeCustomChart,
  getChartDataForSavedChart,
}) {
  if (!savedCharts.length) {
    return (
      <div className="rounded-[2rem] border border-dashed border-slate-300 bg-white/80 p-8 text-center shadow-sm">
        <LayoutDashboard className="mx-auto mb-3 h-9 w-9 text-slate-400" />
        <h3 className="text-lg font-black text-slate-900">
          Custom dashboard is empty
        </h3>
        <p className="mt-2 text-sm text-slate-500">
          Use the Custom Builder on the left to add multiple charts.
        </p>
      </div>
    );
  }

  return (
    <div className="rounded-[2rem] border border-white bg-white/90 p-6 shadow-xl shadow-slate-200/70 backdrop-blur">
      <div className="mb-6 flex items-center gap-3">
        <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-blue-50 text-blue-600">
          <LayoutDashboard className="h-6 w-6" />
        </div>

        <div>
          <h3 className="text-xl font-black text-slate-950">
            Custom Dashboard
          </h3>
          <p className="text-sm text-slate-500">
            {savedCharts.length} chart{savedCharts.length > 1 ? "s" : ""} added
          </p>
        </div>
      </div>

      <div className="grid gap-6 xl:grid-cols-2">
        {savedCharts.map((chart) => (
          <div
            key={chart.id}
            className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm"
          >
            <div className="mb-4 flex items-start justify-between gap-4">
              <div>
                <h4 className="font-black text-slate-950">{chart.title}</h4>
                <p className="mt-1 text-xs font-semibold text-slate-500">
                  {getChartTypeLabel(chart.type)} • {chart.xColumn}
                  {chart.yColumn ? ` • ${chart.yColumn}` : ""}
                </p>
              </div>

              <button
                onClick={() => removeCustomChart(chart.id)}
                className="rounded-xl bg-red-50 p-2 text-red-600 transition hover:bg-red-100"
              >
                <Trash2 className="h-4 w-4" />
              </button>
            </div>

            <SavedChartRenderer
              chart={chart}
              data={getChartDataForSavedChart(chart)}
            />
          </div>
        ))}
      </div>
    </div>
  );
}

function SavedChartRenderer({ chart, data }) {
  if (!data.length) {
    return (
      <div className="rounded-2xl bg-slate-50 p-8 text-center text-sm font-semibold text-slate-500">
        No chart data available.
      </div>
    );
  }

  if (chart.type === "bar") {
    return <VerticalBar data={data} color="#2563eb" />;
  }

  if (chart.type === "horizontal-bar") {
    return <HorizontalBar data={data} color="#2563eb" />;
  }

  if (chart.type === "pie") {
    return <PieVisual data={data} />;
  }

  if (chart.type === "line") {
    return <LineVisual data={data} />;
  }

  return null;
}

function DashboardCard({ title, children }) {
  return (
    <div className="rounded-[2rem] border border-white bg-white/90 p-6 shadow-xl shadow-slate-200/70 backdrop-blur">
      <div className="mb-6 flex items-center gap-3">
        <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-green-50 text-green-600">
          <CheckCircle2 className="h-6 w-6" />
        </div>

        <div>
          <h3 className="text-xl font-black text-slate-950">{title}</h3>
          <p className="text-sm text-slate-500">
            Auto-generated from your uploaded data.
          </p>
        </div>
      </div>

      {children}
    </div>
  );
}

function SelectorGrid({ children }) {
  return <div className="grid gap-4 md:grid-cols-2">{children}</div>;
}

function InsightBox({ text }) {
  return (
    <div className="mt-5 rounded-3xl border border-blue-100 bg-gradient-to-r from-blue-50 to-indigo-50 p-4 text-sm font-bold leading-6 text-blue-900">
      <div className="flex gap-3">
        <Sparkles className="mt-0.5 h-5 w-5 shrink-0 text-blue-600" />
        <span>{text}</span>
      </div>
    </div>
  );
}

function ChartPanel({ title, children }) {
  return (
    <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <h4 className="mb-5 text-sm font-black uppercase tracking-wide text-slate-600">
        {title}
      </h4>
      {children}
    </div>
  );
}

function SelectBox({ label, value, onChange, options }) {
  return (
    <div>
      <label className="mb-2 block text-xs font-black uppercase tracking-wide text-slate-500">
        {label}
      </label>

      <select
        value={value}
        onChange={(event) => onChange(event.target.value)}
        className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-semibold outline-none transition focus:border-blue-500 focus:bg-white focus:ring-4 focus:ring-blue-100"
      >
        <option value="">Select column</option>

        {options.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
    </div>
  );
}

function ColumnGroup({ title, items }) {
  return (
    <div className="mb-5">
      <div className="mb-2 flex items-center justify-between">
        <p className="text-xs font-black uppercase tracking-wide text-slate-500">
          {title}
        </p>
        <span className="rounded-full bg-slate-100 px-2 py-0.5 text-xs font-bold text-slate-500">
          {items.length}
        </span>
      </div>

      {items.length ? (
        <div className="flex flex-wrap gap-2">
          {items.map((item) => (
            <span
              key={item}
              className="rounded-full border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-semibold text-slate-700"
            >
              {item}
            </span>
          ))}
        </div>
      ) : (
        <p className="rounded-2xl bg-slate-50 px-3 py-2 text-sm text-slate-400">
          None detected
        </p>
      )}
    </div>
  );
}

function FeatureCard({ icon, title, description }) {
  return (
    <div className="rounded-3xl border border-slate-200 bg-white p-6 text-left shadow-sm transition hover:-translate-y-1 hover:shadow-xl">
      <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-50 text-blue-600">
        {icon}
      </div>
      <h3 className="font-black text-slate-950">{title}</h3>
      <p className="mt-2 text-sm leading-6 text-slate-600">{description}</p>
    </div>
  );
}

function DataPreview({ columns, data }) {
  return (
    <div className="rounded-[2rem] border border-white bg-white/90 p-6 shadow-xl shadow-slate-200/70 backdrop-blur">
      <div className="mb-5 flex flex-col gap-2 md:flex-row md:items-center md:justify-between">
        <div>
          <h3 className="text-xl font-black text-slate-950">Data Preview</h3>
          <p className="text-sm text-slate-500">
            Showing first 20 filtered records
          </p>
        </div>

        <span className="rounded-full bg-slate-100 px-4 py-2 text-sm font-bold text-slate-600">
          {data.length} rows
        </span>
      </div>

      <div className="overflow-auto rounded-3xl border border-slate-200">
        <table className="min-w-full text-sm">
          <thead className="bg-slate-950">
            <tr>
              {columns.map((column) => (
                <th
                  key={column}
                  className="whitespace-nowrap border-b border-slate-800 px-4 py-4 text-left text-xs font-black uppercase tracking-wide text-white"
                >
                  {column}
                </th>
              ))}
            </tr>
          </thead>

          <tbody className="bg-white">
            {data.slice(0, 20).map((row, index) => (
              <tr
                key={index}
                className="border-b border-slate-100 transition hover:bg-blue-50/40"
              >
                {columns.map((column) => (
                  <td
                    key={column}
                    className="whitespace-nowrap px-4 py-3 font-medium text-slate-600"
                  >
                    {formatCell(row[column])}
                  </td>
                ))}
              </tr>
            ))}

            {data.length === 0 && (
              <tr>
                <td
                  colSpan={columns.length}
                  className="px-4 py-10 text-center font-semibold text-slate-500"
                >
                  No records match the current filters.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function VerticalBar({ data, color }) {
  return (
    <ResponsiveContainer width="100%" height={340}>
      <BarChart data={data}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="name" tick={{ fontSize: 11 }} />
        <YAxis />
        <Tooltip />
        <Bar dataKey="value" fill={color} radius={[8, 8, 0, 0]} />
      </BarChart>
    </ResponsiveContainer>
  );
}

function HorizontalBar({ data, color }) {
  return (
    <ResponsiveContainer width="100%" height={340}>
      <BarChart data={data} layout="vertical">
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis type="number" />
        <YAxis dataKey="name" type="category" width={120} tick={{ fontSize: 11 }} />
        <Tooltip />
        <Bar dataKey="value" fill={color} radius={[0, 8, 8, 0]} />
      </BarChart>
    </ResponsiveContainer>
  );
}

function PieVisual({ data }) {
  return (
    <ResponsiveContainer width="100%" height={340}>
      <PieChart>
        <Pie
          data={data}
          dataKey="value"
          nameKey="name"
          cx="50%"
          cy="50%"
          outerRadius={110}
          label
        >
          {data.map((entry, index) => (
            <Cell key={entry.name} fill={COLORS[index % COLORS.length]} />
          ))}
        </Pie>
        <Tooltip />
        <Legend />
      </PieChart>
    </ResponsiveContainer>
  );
}

function LineVisual({ data }) {
  return (
    <ResponsiveContainer width="100%" height={380}>
      <LineChart data={data}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey="name" tick={{ fontSize: 11 }} />
        <YAxis />
        <Tooltip />
        <Line
          type="monotone"
          dataKey="value"
          stroke="#2563eb"
          strokeWidth={3}
          dot
        />
      </LineChart>
    </ResponsiveContainer>
  );
}

function getChartDataForSavedChart(chart, data) {
  if (chart.type === "pie") {
    return getGroupedCountDataHelper(chart.xColumn, data);
  }

  if (chart.type === "line") {
    return getTimeTrendDataHelper(chart.xColumn, chart.yColumn, data);
  }

  return getGroupedSumDataHelper(chart.xColumn, chart.yColumn, data);
}

function getGroupedSumDataHelper(categoryColumn, numberColumn, data) {
  if (!categoryColumn || !numberColumn) return [];

  const grouped = {};

  data.forEach((row) => {
    const category = row[categoryColumn] || "Missing";
    const value = Number(row[numberColumn]) || 0;

    if (!grouped[category]) grouped[category] = 0;
    grouped[category] += value;
  });

  return Object.entries(grouped)
    .map(([name, value]) => ({
      name,
      value: Number(value.toFixed(2)),
    }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 15);
}

function getGroupedCountDataHelper(categoryColumn, data) {
  if (!categoryColumn) return [];

  const grouped = {};

  data.forEach((row) => {
    const category = row[categoryColumn] || "Missing";

    if (!grouped[category]) grouped[category] = 0;
    grouped[category] += 1;
  });

  return Object.entries(grouped)
    .map(([name, value]) => ({
      name,
      value,
    }))
    .sort((a, b) => b.value - a.value)
    .slice(0, 15);
}

function getTimeTrendDataHelper(dateColumn, numberColumn, data) {
  if (!dateColumn || !numberColumn) return [];

  const grouped = {};

  data.forEach((row) => {
    const rawDate = row[dateColumn];
    const value = Number(row[numberColumn]) || 0;

    let dateKey = "";

    if (rawDate instanceof Date && !isNaN(rawDate)) {
      dateKey = rawDate.toISOString().split("T")[0];
    } else {
      const parsed = new Date(rawDate);
      if (!isNaN(parsed)) {
        dateKey = parsed.toISOString().split("T")[0];
      }
    }

    if (!dateKey) return;

    if (!grouped[dateKey]) grouped[dateKey] = 0;
    grouped[dateKey] += value;
  });

  return Object.entries(grouped)
    .map(([name, value]) => ({
      name,
      value: Number(value.toFixed(2)),
    }))
    .sort((a, b) => new Date(a.name) - new Date(b.name));
}

function getChartTypeLabel(type) {
  if (type === "bar") return "Bar Chart";
  if (type === "horizontal-bar") return "Horizontal Bar";
  if (type === "pie") return "Pie Chart";
  if (type === "line") return "Line Chart";
  return "Chart";
}

function getDashboardIcon(dashboard) {
  if (dashboard.includes("Performance")) {
    return <BarChart3 className="h-4 w-4 text-blue-600" />;
  }

  if (dashboard.includes("Status")) {
    return <PieChartIcon className="h-4 w-4 text-green-600" />;
  }

  if (dashboard.includes("Time")) {
    return <LineChartIcon className="h-4 w-4 text-purple-600" />;
  }

  return <Table className="h-4 w-4 text-slate-600" />;
}

function formatCell(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value.toLocaleDateString();
  }

  if (value === null || value === undefined || value === "") {
    return "-";
  }

  return String(value);
}