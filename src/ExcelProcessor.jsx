import React, { useState } from "react";
import {
  Button,
  Typography,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TablePagination,
} from "@mui/material";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { Buffer } from "buffer";
import { ToastContainer, toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";

function ExcelProcessor() {
  const [excelData, setExcelData] = useState([]);
  const [selectedFile, setSelectedFile] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [page, setPage] = useState(0); // Página actual
  const [rowsPerPage, setRowsPerPage] = useState(5);
  React.useEffect(() => {
    const handleDone = () => {
      toast.success("¡Las fichas de costo se guardaron correctamente!", {
        position: "bottom-right",
        autoClose: 4000,
      });
    };

    window.electron.ipcRenderer.on("save-excel-files-done", handleDone);

    return () => {
      // Limpia al desmontar el componente
      window.electron.ipcRenderer.removeListener(
        "save-excel-files-done",
        handleDone
      );
    };
  }, []);

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    setSelectedFile(file);
  };

  const handleFileUpload = () => {
    if (!selectedFile) {
      alert("Por favor, selecciona un archivo Excel.");
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const binaryString = e.target.result;
      const workbook = XLSX.read(binaryString, { type: "binary" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      if (rawData.length > 0) {
        const extractedHeaders = rawData[0].filter(
          (header) => header !== null && header !== undefined
        );
        setHeaders(extractedHeaders);
        const dataRows = rawData.slice(1).map((row) => {
          const obj = {};
          extractedHeaders.forEach((header, index) => {
            obj[header] = row[index];
          });
          return obj;
        });
        setExcelData(dataRows);
        setPage(0); // Resetear la página al cargar nuevos datos
      } else {
        setExcelData([]);
        setHeaders([]);
        setPage(0);
        alert("El archivo Excel está vacío o no tiene datos.");
      }
    };
    reader.readAsBinaryString(selectedFile);
  };

  const handleExport = async () => {
    let gastoMaterial = 0;
    let costoTotal = 0;
    let gastosGeneralesAdmin = 0;
    let salariosGastosGenAdmin = 0;
    let gastosDistribVenta = 0;
    let salariosDistribVenta = 0;
    let gastosFinancieros = 0;
    let totalGastos = 0;
    let totalCostos = 0;
    let tasaUtilidad = 0.15;
    let utilidad = 0;
    let precio = 0;
    let precioUnitario = 0;
    let archivosExcelArray = [];

    for (const row of excelData) {
      gastoMaterial = +row["Costo Base"];
      costoTotal = gastoMaterial;
      gastosGeneralesAdmin = gastoMaterial * (1.25 * 0.104108830229265);
      salariosGastosGenAdmin =
        gastoMaterial * (1.25 * (0.38 * 0.579778660977904));
      gastosDistribVenta = gastoMaterial * (1.25 * 0.313202984792803);
      salariosDistribVenta =
        gastoMaterial * (1.25 * (0.62 * 0.579778660977904));
      gastosFinancieros = gastoMaterial * ((0.3 / 100) * 1.25);
      totalGastos =
        gastosGeneralesAdmin + gastosDistribVenta + gastosFinancieros;
      totalCostos = costoTotal + totalGastos;
      utilidad = tasaUtilidad * gastoMaterial;
      precio = totalCostos + utilidad;
      precioUnitario = precio / 1;

      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("FICHA");

      //Anchos de columna
      sheet.columns = [
        { width: 15 }, // A
        { width: 15 }, // B
        { width: 15 }, // C
        { width: 25 }, // D en adelante si necesitas más
        { width: 15 },
        { width: 15 },
        { width: 15 },
      ];

      // Título principal
      sheet.mergeCells("C1:G1");
      sheet.getCell("C1").value = "MINISTERIO DE FINANZAS Y PRECIOS";
      sheet.getCell("C1").alignment = { horizontal: "center" };
      sheet.getCell("C1").font = { bold: true };

      sheet.mergeCells("C2:G2");
      sheet.getCell("C2").value = "PRECIOS Y TARIFAS";
      sheet.getCell("C2").alignment = { horizontal: "center" };
      sheet.getCell("C2").font = { bold: true };

      // Subtítulos
      sheet.mergeCells("C3:D3");
      sheet.getCell("C3").value = "Producto o Servicio:";
      sheet.getCell("C3").font = { bold: true };
      sheet.getCell("C3").alignment = { horizontal: "center" };
      sheet.mergeCells("E3:G3");
      sheet.getCell("E3").value = row["Productos"];
      sheet.getCell("E3").font = { bold: true };
      sheet.getCell("E3").alignment = { horizontal: "center" };

      sheet.mergeCells("C4:D4");
      sheet.getCell("C4").value = "Código Prod. o Serv.:";
      sheet.getCell("C4").font = { bold: true };

      sheet.getCell("F4").value = "UM:";
      sheet.getCell("F4").font = { bold: true };
      sheet.mergeCells("G4:H4");
      sheet.getCell("G4").value = "Nivel de Producción:";
      sheet.getCell("G4").font = { bold: true };
      sheet.getCell("G4").alignment = { horizontal: "center" };

      sheet.getCell("F5").value = "UNO";
      sheet.getCell("F5").font = { bold: true };
      sheet.mergeCells("G5:H5");
      sheet.getCell("G5").value = 1;
      sheet.getCell("G5").font = { bold: true };
      sheet.getCell("G5").alignment = { horizontal: "center" };

      // Encabezados de tabla
      sheet.mergeCells("C6:E6");
      sheet.getCell("C6").value = "CONCEPTO";
      sheet.mergeCells("F6:G6");
      sheet.getCell("F6").value = "COSTO BASE";
      ["C6", "D6", "E6", "F6", "G6"].forEach((cell) => {
        sheet.getCell(cell).font = { bold: true };
        sheet.getCell(cell).alignment = { horizontal: "center" };
        sheet.getCell(cell).border = getBorder();
      });

      // Datos (puedes ajustar los valores como los necesites)
      const rows = [
        ["Gasto Material", gastoMaterial.toFixed(2)],
        [
          "    De ello: Insumos (Materias primas y materiales)",
          gastoMaterial.toFixed(2),
        ],
        ["    Combustibles y lubricantes", 0.0],
        ["    Energía", 0.0],
        ["    Agua", 0.0],
        ["Salario Directo o retribución directa", ""],
        ["Otros Gastos Directos (Desglosar)", ""],
        ["Gastos asociados a la producción", ""],
        ["    De ello, salarios", ""],
        ["COSTO TOTAL", costoTotal.toFixed(2)],
        [
          "Gastos Generales y de Administración",
          gastosGeneralesAdmin.toFixed(2),
        ],
        ["    De ello, salarios", salariosGastosGenAdmin.toFixed(2)],
        ["Gastos de Distribución y Venta", gastosDistribVenta.toFixed(2)],
        ["    De ello, salarios", salariosDistribVenta.toFixed(2)],
        ["Gastos Financieros", gastosFinancieros.toFixed(2)],
        ["Gastos por Financiamiento entregado a la OSDE", 0],
        ["Gastos Tributarios (Contribución a la Seguridad)", 0],
        ["TOTAL DE GASTOS", totalGastos.toFixed(2)],
        ["TOTAL DE COSTOS Y GASTOS", totalCostos.toFixed(2)],
        ["Tasa de utilidad", "15%"],
        ["Utilidad", utilidad.toFixed(2)],
        ["PRECIO O TARIFA", precio.toFixed(2)],
        ["PRECIO O TARIFA UNITARIO AJUSTADO", precioUnitario.toFixed(2)],
        ["Datos sobre precios de referencia", ""],
        // ["Elaborado", "Firma:", "Cargo:"],
        // ["Aprobado", "Firma:", "Cargo:"],
      ];

      // Insertar datos con fusiones y estilos (sin bordes)
      rows.forEach((row, index) => {
        const [concepto, costo] = row;
        const rowIndex = index + 7;
        const esTitulo = costo === null || costo === "" || costo === undefined;

        // FUSIÓN CONCEPTO C:E
        sheet.mergeCells(`C${rowIndex}:E${rowIndex}`);
        const cellConcepto = sheet.getCell(`C${rowIndex}`);
        cellConcepto.value = concepto;
        cellConcepto.alignment = {
          vertical: "middle",
          horizontal: "left",
          wrapText: true,
        };
        if (esTitulo) {
          cellConcepto.font = { bold: true };
        }

        // Bordes del bloque C:E
        ["C", "D", "E"].forEach((col, idx, arr) => {
          const cell = sheet.getCell(`${col}${rowIndex}`);
          const border = {};
          border.top = { style: "thin" };
          border.bottom = { style: "thin" };
          if (idx === 0) border.left = { style: "thin" }; // C
          if (idx === arr.length - 1) border.right = { style: "thin" }; // E
          cell.border = border;
        });

        // FUSIÓN COSTO BASE F:G
        sheet.mergeCells(`F${rowIndex}:G${rowIndex}`);
        const cellCosto = sheet.getCell(`F${rowIndex}`);
        if (!esTitulo) {
          cellCosto.value = costo;
        }
        cellCosto.alignment = {
          vertical: "middle",
          horizontal: "center",
          wrapText: true,
        };

        // Bordes del bloque F:G
        ["F", "G"].forEach((col, idx, arr) => {
          const cell = sheet.getCell(`${col}${rowIndex}`);
          const border = {};
          border.top = { style: "thin" };
          border.bottom = { style: "thin" };
          if (idx === 0) border.left = { style: "thin" }; // F
          if (idx === arr.length - 1) border.right = { style: "thin" }; // G
          cell.border = border;
        });
      });

      // Ahora aplicar bordes **solo exteriores** al rango completo de la tabla:

      const startRow = 6; // Primera fila de encabezados
      const endRow = rows.length + 6; // Última fila de datos
      const startCol = 3; // Columna C
      const endCol = 7; // Columna G

      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          const cellAddress = `${String.fromCharCode(64 + col)}${row}`;
          const cell = sheet.getCell(cellAddress);

          const border = {};
          if (row === startRow) border.top = { style: "thin" };
          if (row === endRow) border.bottom = { style: "thin" };
          if (col === startCol) border.left = { style: "thin" };
          if (col === endCol) border.right = { style: "thin" };

          // cell.border = border;

          cell.border = {
            ...cell.border, // 👈 mantiene bordes existentes
            ...(row === startRow ? { top: { style: "thin" } } : {}),
            ...(row === endRow ? { bottom: { style: "thin" } } : {}),
            ...(col === startCol ? { left: { style: "thin" } } : {}),
            ...(col === endCol ? { right: { style: "thin" } } : {}),
          };
        }
      }

      const rowOffset = rows.length + 4;

      // Fila en blanco
      // sheet.getRow(rowOffset + 3).values = [];

      // Fila "RANGO DE PRECIOS" centrado
      sheet.mergeCells(`C${rowOffset + 4}:D${rowOffset + 4}`);
      sheet.getCell(`C${rowOffset + 4}`).value = "RANGO DE PRECIOS";
      sheet.getCell(`C${rowOffset + 4}`).alignment = { horizontal: "center" };
      sheet.getCell(`C${rowOffset + 4}`).font = { bold: true, underline: true };

      // Fila "MÍNIMO"
      sheet.getCell(`C${rowOffset + 5}`).value = "MÍNIMO";
      sheet.getCell(`C${rowOffset + 5}`).font = { bold: true };
      sheet.getCell(`E${rowOffset + 5}`).value = (precio * (100 - 10)) / 100;
      sheet.getCell(`E${rowOffset + 5}`).numFmt = "0.00";

      // Fila "MÁXIMO"
      sheet.getCell(`C${rowOffset + 6}`).value = "MÁXIMO";
      sheet.getCell(`C${rowOffset + 6}`).font = { bold: true };
      sheet.getCell(`E${rowOffset + 6}`).value = (precio * (100 + 10)) / 100;
      sheet.getCell(`E${rowOffset + 6}`).numFmt = "0.00";

      sheet.getCell("C7").font = { bold: true };
      sheet.getCell("C12").font = { bold: true };
      sheet.getCell("C13").font = { bold: true };
      sheet.getCell("C14").font = { bold: true };
      sheet.getCell("C16").font = { bold: true };
      sheet.getCell("C17").font = { bold: true };
      sheet.getCell("C19").font = { bold: true };
      sheet.getCell("C21").font = { bold: true };
      sheet.getCell("C22").font = { bold: true };
      sheet.getCell("C23").font = { bold: true };
      sheet.getCell("C24").font = { bold: true };
      sheet.getCell("C25").font = { bold: true };
      sheet.getCell("C26").font = { bold: true };
      sheet.getCell("C27").font = { bold: true };
      sheet.getCell("C28").font = { bold: true };
      sheet.getCell("C29").font = { bold: true };
      sheet.getCell("C30").font = { bold: true };

      sheet.getCell("D7").font = { bold: true };
      sheet.getCell("D12").font = { bold: true };
      sheet.getCell("D13").font = { bold: true };
      sheet.getCell("D14").font = { bold: true };
      sheet.getCell("D16").font = { bold: true };
      sheet.getCell("D17").font = { bold: true };
      sheet.getCell("D19").font = { bold: true };
      sheet.getCell("D21").font = { bold: true };
      sheet.getCell("D22").font = { bold: true };
      sheet.getCell("D23").font = { bold: true };
      sheet.getCell("D24").font = { bold: true };
      sheet.getCell("D25").font = { bold: true };
      sheet.getCell("D26").font = { bold: true };
      sheet.getCell("D27").font = { bold: true };
      sheet.getCell("D28").font = { bold: true };
      sheet.getCell("D29").font = { bold: true };
      sheet.getCell("D30").font = { bold: true };

      sheet.getCell("E3").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFF00" }, // Amarillo
      };

      // Exportar
      const buffer = await workbook.xlsx.writeBuffer();
      archivosExcelArray.push({
        nombre: `${row["Productos"]}.xlsx`,
        buffer: Buffer.from(buffer), // Muy importante: pasar a Node Buffer
      });
    }

    window.electron.ipcRenderer.send("save-excel-files", archivosExcelArray);
  };

  // Función de bordes
  function getBorder() {
    return {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  }

  const handleChangePage = (event, newPage) => {
    setPage(newPage);
  };

  const handleChangeRowsPerPage = (event) => {
    setRowsPerPage(parseInt(event.target.value, 10));
    setPage(0); // Resetear la página al cambiar la cantidad de filas por página
  };

  const emptyRows =
    rowsPerPage - Math.min(rowsPerPage, excelData.length - page * rowsPerPage);

  return (
    <Paper
      elevation={3}
      style={{
        padding: "20px",
        maxWidth: "800px",
        margin: "20px auto",
        width: "900px",
      }}
    >
      <Typography variant="h5" gutterBottom>
        Procesador de datos para Fichas de costos
      </Typography>
      <input
        type="file"
        accept=".xlsx, .csv"
        onChange={handleFileChange}
        style={{ marginBottom: "10px" }}
      />
      <Button
        variant="contained"
        color="primary"
        onClick={handleFileUpload}
        disabled={!selectedFile}
        style={{ marginBottom: "20px" }}
      >
        Cargar Excel
      </Button>

      {excelData.length > 0 && (
        <>
          <Typography variant="h6" gutterBottom>
            Datos Cargados
          </Typography>
          <TableContainer component={Paper} style={{ marginBottom: "20px" }}>
            <Table>
              <TableHead>
                <TableRow>
                  {headers.map((header) => (
                    <TableCell key={header} sx={{ fontWeight: "bold" }}>
                      {header}
                    </TableCell>
                  ))}
                </TableRow>
              </TableHead>
              <TableBody>
                {(rowsPerPage > 0
                  ? excelData.slice(
                      page * rowsPerPage,
                      page * rowsPerPage + rowsPerPage
                    )
                  : excelData
                ).map((row, index) => (
                  <TableRow key={index}>
                    {headers.map((header) => (
                      <TableCell key={`${index}-${header}`}>
                        {row[header]}
                      </TableCell>
                    ))}
                  </TableRow>
                ))}

                {emptyRows > 0 && (
                  <TableRow style={{ height: 53 * emptyRows }}>
                    <TableCell colSpan={headers.length} />
                  </TableRow>
                )}
              </TableBody>
            </Table>
          </TableContainer>

          <TablePagination
            rowsPerPageOptions={[5, 10, 25]}
            component="div"
            count={excelData.length}
            rowsPerPage={rowsPerPage}
            page={page}
            onPageChange={handleChangePage}
            onRowsPerPageChange={handleChangeRowsPerPage}
            labelRowsPerPage="Filas por página:"
            labelDisplayedRows={({ from, to, count }) =>
              `de ${from}-${to} de ${count}`
            }
          />
          <Button variant="contained" color="success" onClick={handleExport}>
            Exportar a Excel
          </Button>
        </>
      )}
      <ToastContainer />
    </Paper>
  );
}

export default ExcelProcessor;
