import "./App.css";

// Require library
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  BorderStyle,
  HeadingLevel,
  AlignmentType,
  TextRun,
  LevelFormat,
  UnderlineType,
  convertInchesToTwip,
  WidthType,
} from "docx";
import { saveAs } from "file-saver";

const App = () => {
  const generateExcel = () => {
    const doc = new Document({
      creator: "Clippy",
      title: "Sample Document",
      description: "A brief example of using docx",
      styles: {
        paragraphStyles: [
          {
            id: "title",
            name: "Title",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 36,
              bold: true,
              font: "Calibri",
              color: "black",
            },
          },
          {
            id: "law",
            name: "Law",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 18,
              font: "Calibri",
              color: "black",
            },
          },
          {
            id: "company",
            name: "Company",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 24,
              bold: true,
              font: "Calibri",
              color: "#77b516",
            },
          },
          {
            id: "address",
            name: "Address",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 22,
              bold: true,
              font: "Calibri",
              color: "#77b516",
            },
          },
          {
            id: "date",
            name: "Date",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 36,
              font: "Calibri",
              color: "black",
            },
          },
          {
            id: "counter",
            name: "Counter",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 36,
              font: "Calibri",
              color: "red",
              bold: true,
            },
          },
          {
            id: "textCounter",
            name: "TextCounter",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 16,
              font: "Calibri",
              color: "red",
            },
          },
          {
            id: "tableHead",
            name: "TableHead",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 16,
              font: "Calibri",
            },
          },
          {
            id: "tableHeadLog",
            name: "TableHead",
            basedOn: "Normal",
            next: "Normal",
            run: {
              size: 22,
              font: "Calibri",
            },
          },
        ],
      },
      numbering: {
        config: [
          {
            reference: "my-crazy-numbering",
            levels: [
              {
                level: 0,
                format: LevelFormat.LOWER_LETTER,
                text: "%1)",
                alignment: AlignmentType.LEFT,
              },
            ],
          },
        ],
      },
      sections: [
        {
          children: [
            new Paragraph({
              text:
                "Documento soporte de costos y gastos en operaciones con no obligados a expedir factura o documento equivalente",
              style: "title",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text:
                "Artículo 1.6.1.4.12 Decreto Único reglamentario en materia tributaria 1625 de 2016 - Sustituido por el Decreto 358 de 2020",
              style: "law",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph(""),
            new Paragraph({
              text: "OIKONOMOS S.A.S. Soluciones financieras.",
              style: "company",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: "NIT: 901355211-1",
              style: "company",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
              text: "Calle 26a #13-97, oficina 1204",
              style: "address",
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph(""),
            new Table({
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: {
                        size: 6005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: `Fecha de operación: ${new Date().getDate()}/${
                            new Date().getMonth() + 1
                          }/${new Date().getFullYear()}`,
                          style: "date",
                        }),
                      ],
                      borders: {
                        top: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        right: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        bottom: {
                          style: BorderStyle.THICK_THIN_MEDIUM_GAP,
                          size: 0,
                          color: "white",
                        },
                        left: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                      },
                    }),
                    new TableCell({
                      width: {
                        size: 3005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "No. XXX",
                          style: "counter",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                      borders: {
                        top: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        right: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        bottom: {
                          style: BorderStyle.THICK_THIN_MEDIUM_GAP,
                          size: 0,
                          color: "white",
                        },
                        left: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                      },
                    }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({
                      width: {
                        size: 6005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "",
                          style: "textCounter",
                        }),
                      ],
                      borders: {
                        top: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        right: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        bottom: {
                          style: BorderStyle.THICK_THIN_MEDIUM_GAP,
                          size: 0,
                          color: "white",
                        },
                        left: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                      },
                    }),
                    new TableCell({
                      width: {
                        size: 3005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "Autorización de numeración Dian No",
                          style: "textCounter",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                      borders: {
                        top: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        right: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        bottom: {
                          style: BorderStyle.THICK_THIN_MEDIUM_GAP,
                          size: 0,
                          color: "white",
                        },
                        left: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                      },
                    }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({
                      width: {
                        size: 6005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "",
                          style: "textCounter",
                        }),
                      ],
                      borders: {
                        top: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        right: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        bottom: {
                          style: BorderStyle.THICK_THIN_MEDIUM_GAP,
                          size: 0,
                          color: "white",
                        },
                        left: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                      },
                    }),
                    new TableCell({
                      width: {
                        size: 3005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "XXXXXXXXXX  desde  DSOP 1 hasta la 250",
                          style: "textCounter",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                      borders: {
                        top: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        right: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                        bottom: {
                          style: BorderStyle.THICK_THIN_MEDIUM_GAP,
                          size: 0,
                          color: "white",
                        },
                        left: {
                          style: BorderStyle.DASH_DOT_STROKED,
                          size: 0,
                          color: "white",
                        },
                      },
                    }),
                  ],
                }),
              ],
            }),
            new Paragraph(""),
            new Table({
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      width: {
                        size: 7005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "Vendedor o quien presta el servicio:",
                          style: "tableHead",
                        }),
                      ],
                    }),
                    new TableCell({
                      width: {
                        size: 2005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "Nit:",
                          style: "tableHead",
                        }),
                      ],
                    }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({
                      width: {
                        size: 7005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "Johan David Pedroza Plazas",
                          style: "tableHeadLog",
                        }),
                      ],
                    }),
                    new TableCell({
                      width: {
                        size: 2005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "1010237909",
                          style: "tableHeadLog",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({
                      width: {
                        size: 7005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "Direccion:",
                          style: "tableHead",
                        }),
                      ],
                    }),
                    new TableCell({
                      width: {
                        size: 2005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "Telefonos:",
                          style: "tableHead",
                        }),
                      ],
                    }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({
                      width: {
                        size: 7005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "Calle 70 #48-26 sur",
                          style: "tableHeadLog",
                        }),
                      ],
                    }),
                    new TableCell({
                      width: {
                        size: 2005,
                        type: WidthType.DXA,
                      },
                      children: [
                        new Paragraph({
                          text: "3209424973",
                          style: "tableHeadLog",
                          alignment: AlignmentType.CENTER,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        },
      ],
    });

    Packer.toBlob(doc).then((blod) => {
      saveAs(blod, "firstDocument");
    });
  };

  return (
    <div className="App">
      <div
        onClick={generateExcel}
        style={{
          color: "white",
          border: "solid 2px black",
          background: "red",
          padding: 5,
          width: "20%",
          margin: "25px auto 0 auto",
        }}
      >
        Generar documento
      </div>
    </div>
  );
};

export default App;
