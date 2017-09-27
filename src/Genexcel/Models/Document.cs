using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;

namespace Genexcel {
	public class Document {
		const int _sheetNameLengthLimit = 31;
		static uint _idGen = 1;
		HashSet<Models.Sheet> _sheets = new HashSet<Models.Sheet>();
		public Document() {
			this._sheets.Add(new Models.Sheet(this, "Plan1"));
		}

		public Models.Sheet AddSheet(string title) {
			var sheet = new Models.Sheet(this, title);
			this._sheets.Add(sheet);
			return sheet;
			//return AddSheet(new Models.Sheet(this, title));
		}
		//public Models.Sheet AddSheet(Models.Sheet sheet) {
		//	this._sheets.Add(sheet);
		//	return sheet;
		//}

		public IEnumerable<Models.Sheet> GetSheets() {
			return _sheets.ToList();
		}
		
		//void InitBasicStylePart(WorkbookPart workbookPart) {
		//	WorkbookStylesPart stylesPart;
		//	if (!workbookPart.GetPartsOfType<WorkbookStylesPart>().Any()) {
		//		stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
		//		stylesPart.Stylesheet = new Stylesheet();
		//		var stylesheet = stylesPart.Stylesheet;
		//		stylesheet.Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts(
		//			new Font(
		//				new FontSize() { Val = 11 },
		//				new Color() { Theme = 1 }
		//			)
		//		);

		//		stylesheet.CellStyleFormats = new CellStyleFormats();
		//		stylesheet.CellStyleFormats.Append(new CellFormat());
		//		stylesheet.CellFormats = new CellFormats();
		//	}
		//	var cellFormat = stylesheet.CellFormats.Elements<CellFormat>().FirstOrDefault(cf => cf.FormatId == cellStyle.FormatId)
		//	   ?? stylesheet.CellFormats.AppendChild(new CellFormat() {
		//		   FormatId = cellStyle.FormatId,
		//	   });


		//	if (stylesheet.CellStyles == null) {
		//		stylesheet.CellStyles = new CellStyles();
		//	}
		//	var cellStyles = stylesheet.CellStyles;
		//	var cellStyle = cellStyles.Elements<CellStyle>().FirstOrDefault(cs => cs.Name == "Hyperlink")
		//		?? cellStyles.AppendChild(new CellStyle() {
		//			Name = "Hyperlink",
		//			BuiltinId = 8,
		//			FormatId = 0 //index 0 from cellstyleformats
		//					});







		//}

		void Save(SpreadsheetDocument spreadsheetDocument) {
			//Create workbook parts
			var workbookPart = spreadsheetDocument.AddWorkbookPart();
			//Sets workbook
			var workbook = new Workbook();
			workbookPart.Workbook = workbook;

			//Set theme
			using (var stream = GetType().GetTypeInfo().Assembly.GetManifestResourceStream("Genexcel.Resources.Office.theme1.xml")) {
				using (var reader = new StreamReader(stream)) {
					var xml = reader.ReadToEnd();
					workbookPart.AddNewPart<ThemePart>();
					workbookPart.ThemePart.Theme = new Theme(xml);
				}
			}

			//Set styles
			using (var stream = GetType().GetTypeInfo().Assembly.GetManifestResourceStream("Genexcel.Resources.Office.styles.xml")) {
				using (var reader = new StreamReader(stream)) {
					var xml = reader.ReadToEnd();
					workbookPart.AddNewPart<WorkbookStylesPart>();
					workbookPart.WorkbookStylesPart.Stylesheet = new Stylesheet(xml);
				}
			}


			//Adiciona lista sheets
			var sheets = workbook.AppendChild(new Sheets());

			//Adiciona as planilhas ao workbook
			uint sheetId = 1;

			foreach (var s in _sheets) {
				//Criar worksheet part no workbookpart
				var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
				var worksheet = new Worksheet();
				//Columns
				if (s.HasCustomColumn) {
					var columns = new Columns();
					worksheet.Append(columns);
					Column currentColElement = null;
					Models.Column currentColModel = new Models.Column();//fake current
					for(uint i = 0; i < s.Columns.Length; i++) {
						var col = s.Columns[i];
						if(col == currentColModel) {
							currentColElement.Max = i + 1;
						} else {
							currentColElement = new Column() {
								//Style = 1,
								Min = i + 1,
								Max = i + 1,
								//CustomWidth = false,
								Width = Models.Column.DEFAULT_WIDTH
							};
							if(col != null) {
								currentColElement.CustomWidth = true;
								currentColElement.Width = col.Width;
							}
							columns.Append(currentColElement);
						}
						currentColModel = col;
					}

					//Column currentColumnElement = null;
					//Models.Column currentColumnModel = null;
					//foreach (var col in s.Columns) {
					//	Column colElement;
					//	if(col == currentColumnModel) { colElement = currentColumnElement }
					//	if(col == null) {

					//	}
					//	columns.Append(new Column() {
					//		Min = (uint)col.Min,
					//		Max = (uint)col.Max,
					//		Width = col.Width,
					//		Style = 1,
					//		CustomWidth = true
					//	});
					//}
				}
				var sheetData = new SheetData();
				worksheet.Append(sheetData);
				worksheetPart.Worksheet = worksheet;
				
				var name = s.Name ?? "Plan";
				name = name.Length > _sheetNameLengthLimit ?
					name.Substring(0, _sheetNameLengthLimit) :
					name;
				// Append a new worksheet and associate it with the workbook.
				var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() {
					Id = workbookPart.GetIdOfPart(worksheetPart),
					SheetId = sheetId++,
					Name = name
				};
				sheets.Append(sheet);

				


				foreach (var c in s.GetCells()) {
					// Insert cell A1 into the new worksheet.
					var cell = InsertCellInWorksheet(ColTranslate(c.Col), (uint)c.Row, worksheetPart);
					var value = c.Value;
					if (value is string) {
						//Inicializa sharedStrings se necessário
						SharedStringTablePart shareStringPart;
						if (workbookPart.GetPartsOfType<SharedStringTablePart>().Any()) {
							shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
						} else {
							shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
						}

						// Insert the text into the SharedStringTablePart.
						int index = InsertSharedStringItem(value.ToString(), shareStringPart);

						cell.CellValue = new CellValue(index.ToString());
						cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
					} else if (value  is int
							 || value is decimal
							 || value is long
							 || value is short
							 || value is double
							 || value is float
							 || value is byte) {
						var toString = value.GetType().GetTypeInfo()
							.GetDeclaredMethods("ToString")
							.First(m => m.GetParameters().Any(p => p.ParameterType == typeof(IFormatProvider)));//.GetMethod("ToString", new Type[] { typeof(CultureInfo) }).GetMethodInfo();
						var formattedValue = toString.Invoke(value, new object[] { new CultureInfo("en-US") }).ToString();
						cell.CellValue = new CellValue(formattedValue);
						cell.DataType = new EnumValue<CellValues>(CellValues.Number);
					}

					if (!string.IsNullOrWhiteSpace(c.Hyperlink)) {
						var rId = $"r{Guid.NewGuid().ToString()}";
						//if (workbookPart.GetPartsOfType<Relat>().Any()) {
						//	shareStringPart = worksheet.getp.GetPartsOfType<SharedStringTablePart>().First();
						//} else {
						//	shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
						//}
						var rel = worksheetPart.AddHyperlinkRelationship(new Uri(c.Hyperlink), true, rId);
						var hyperlinks = worksheet.Elements<Hyperlinks>().FirstOrDefault();
						if(hyperlinks == null) { hyperlinks = worksheet.AppendChild(new Hyperlinks()); }
						hyperlinks.Append(new DocumentFormat.OpenXml.Spreadsheet.Hyperlink() {
							Reference = cell.CellReference,
							Id = rId
						});
						
						cell.StyleIndex = 2;//Hyperlink, should be an enum
					}
				}

			
				//Charts
				foreach (var ch in s.Charts) {
					//https://msdn.microsoft.com/en-us/library/office/cc820055.aspx#How the Sample Code Works
					// Add a new drawing to the worksheet.
					var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
					worksheetPart.Worksheet.Append(new Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) });
					worksheetPart.Worksheet.Save();
					var chartPart = drawingsPart.AddNewPart<ChartPart>();
					var chartSpace = new ChartSpace();
					chartPart.ChartSpace = chartSpace;
					chartSpace.Append(new Date1904() { Val = false });
					chartSpace.Append(new EditingLanguage() { Val = "en-US" });
					chartSpace.Append(new RoundedCorners() { Val = false });
					var chart = chartSpace.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Chart());
					//chartSpace.Append(new ChartShapeProperties(
					//						new SolidFill(
					//							new SchemeColor() { Val = SchemeColorValues.Background1 }
					//						),
					//						new DocumentFormat.OpenXml.Drawing.Outline(
					//							new SolidFill(
					//								new SchemeColor(
					//									new LuminanceModulation() { Val = 15000 },
					//									new LuminanceOffset() { Val = 85000 }
					//								) { Val = SchemeColorValues.Text1 }
					//							)
					//						) {
					//							Width = 9525,
					//							CapType = LineCapValues.Flat,
					//							CompoundLineType = CompoundLineValues.Single,
					//							Alignment = PenAlignmentValues.Center
					//						}
					//				));

					//Dont know
					chart.AppendChild(new Title(
						new Overlay() { Val = false },
						new ChartShapeProperties(
							new NoFill(),
							new DocumentFormat.OpenXml.Drawing.Outline(new NoFill()),
							new EffectList()
						),
							new DocumentFormat.OpenXml.Drawing.Charts.TextProperties(
										new BodyProperties() {
											Rotation = 0,
											UseParagraphSpacing = true,
											VerticalOverflow = TextVerticalOverflowValues.Ellipsis,
											Vertical = TextVerticalValues.Horizontal,
											Wrap = TextWrappingValues.Square,
											Anchor = TextAnchoringTypeValues.Center,
											AnchorCenter = true,
										},
										new Paragraph(
											new ParagraphProperties(
												new DefaultRunProperties(
													new SolidFill(
														new SchemeColor(
															new LuminanceModulation() { Val = 65000 },
															new LuminanceOffset() { Val = 35000 }
														) {
															Val = SchemeColorValues.Text1
														}
													),
													new LatinFont() { Typeface = "+mn-lt" },
													new EastAsianFont() { Typeface = "+mn-ea" },
													new ComplexScriptFont() { Typeface = "+mn-cs" }
												) {
													FontSize = 1400,
													Bold = false,
													Italic = false,
													Underline = TextUnderlineValues.None,
													Strike = TextStrikeValues.NoStrike,
													Kerning = 1200,
													Baseline = 0
												}
											)
										)
									)
					));

					//Allow showing title on top
					chart.AppendChild(new AutoTitleDeleted() { Val = false });

					//Create plot area
					var plotArea = chart.AppendChild(new PlotArea());

					var layout = plotArea.AppendChild(new Layout());
					if (ch is Models.AreaChart || ch is Models.BarChart) {

						#region init chart
						var chObject = ch as Models.Chart;
						OpenXmlCompositeElement chartElement;
						if (ch is Models.AreaChart) {
							chartElement = plotArea.AppendChild(
									//Dont know what extensions are for
									new AreaChart(new Grouping() { Val = GroupingValues.Standard })
								);
						} else {
							chartElement = plotArea.AppendChild(
									//Dont know what extensions are for
									new BarChart(
										new BarDirection() {
											Val = BarDirectionValues.Column
										},
										new BarGrouping() {
										Val = BarGroupingValues.Clustered
									})
								);
						}
						chartElement.AppendChild(new VaryColors() { Val = false });
						#endregion

						#region data
						foreach (var dts in chObject.Data.Datasets) {
							var index = (uint)chObject.Data.Datasets.IndexOf(dts);
							OpenXmlCompositeElement chartSeries;
							if (ch is Models.AreaChart) {
								chartSeries = chartElement.AppendChild(new AreaChartSeries());
							} else {
								chartSeries = chartElement.AppendChild(new BarChartSeries());
							}
							chartSeries.Append(
								new Index() { Val = index },
								new Order() { Val = index },
								new SeriesText() { NumericValue = new NumericValue(dts.Title) },
								new ChartShapeProperties(
									new SolidFill(new SchemeColor() { Val = SchemeColorValues.Accent1 }),
									new DocumentFormat.OpenXml.Drawing.Outline(new NoFill()),
									new EffectList()
								)
							);

							if(ch is Models.BarChart) {
								chartSeries.Append(new InvertIfNegative() { Val = false });
							}
							

							//Eixo x (labels)
							var categoryAxisData = chartSeries.AppendChild(new CategoryAxisData());
							var strLit = categoryAxisData.AppendChild(new StringLiteral());
							strLit.Append(new PointCount() { Val = (uint)chObject.Data.Labels.Count });
							foreach (var lbl in chObject.Data.Labels) {
								strLit.AppendChild(new StringPoint() { Index = (uint)chObject.Data.Labels.IndexOf(lbl) })
									.Append(new NumericValue(lbl));
							}

							var values = chartSeries.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values());
							var numLit = values.AppendChild(new NumberLiteral());
							numLit.Append(new FormatCode("General"));
							numLit.Append(new PointCount() { Val = (uint)chObject.Data.Labels.Count });
							foreach (var lbl in chObject.Data.Labels) {
								var lblIndex = chObject.Data.Labels.IndexOf(lbl);
								var val = dts.Data.Count > lblIndex ? dts.Data[lblIndex] : 0;
								numLit.AppendChild(new NumericPoint() { Index = (uint)chObject.Data.Labels.IndexOf(lbl) })
									.Append(new NumericValue(val.ToString()));
							}
							//			numLit.AppendChild(new NumericPoint() { Index = new UInt32Value(0u) })
							//	.Append
							//(new NumericValue("28"));
						}
						#endregion

						#region options?
						//Not required for a valid xlsx
						chartElement
							.AppendChild(
								new DataLabels(
									new ShowLegendKey() { Val = false },
									new ShowValue() { Val = false },
									new ShowCategoryName() { Val = false },
									new ShowSeriesName() { Val = false },
									new ShowPercent() { Val = false },
									new ShowBubbleSize() { Val = false }
								)
							);

						if(ch is Models.BarChart) {
							chartElement.Append(new GapWidth() { Val = 219 });
							chartElement.Append(new Overlap() { Val = -27 });
						}
						#endregion

						#region Axis
						chartElement.Append(new AxisId() { Val = 48650112u });
						chartElement.Append(new AxisId() { Val = 48672768u });

						// Add the Category Axis.
						var catAx = plotArea
							.AppendChild(
								new CategoryAxis(
									new AxisId() { Val = 48650112u },
									new Scaling(
										new Orientation() {
											Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax
										}
									),
									new Delete() { Val = false },
									new AxisPosition() { Val = AxisPositionValues.Bottom },
									new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() {
										FormatCode = "General",
										SourceLinked = true
									},
									new MajorTickMark() { Val = ch is Models.AreaChart ? TickMarkValues.Outside : TickMarkValues.None },
									new MinorTickMark() { Val = TickMarkValues.None },
									new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
									new ChartShapeProperties(
											new NoFill(),
											new DocumentFormat.OpenXml.Drawing.Outline(
												new SolidFill(
													new SchemeColor(
														new LuminanceModulation() { Val = 15000 },
														new LuminanceOffset() { Val = 85000 }
													) { Val = SchemeColorValues.Text1 }
												)
											) {
												Width = 9525,
												CapType = LineCapValues.Flat,
												CompoundLineType = CompoundLineValues.Single,
												Alignment = PenAlignmentValues.Center
											}
									),
									new DocumentFormat.OpenXml.Drawing.Charts.TextProperties(
										new BodyProperties() {
											Rotation = -60000000,
											UseParagraphSpacing = true,
											VerticalOverflow = TextVerticalOverflowValues.Ellipsis,
											Vertical = TextVerticalValues.Horizontal,
											Wrap = TextWrappingValues.Square,
											Anchor = TextAnchoringTypeValues.Center,
											AnchorCenter = true,
										},
										new Paragraph(
											new ParagraphProperties(
												new DefaultRunProperties(
													new SolidFill(
														new SchemeColor(
															new LuminanceModulation() { Val = 65000 },
															new LuminanceOffset() { Val = 35000 }
														) {
															Val = SchemeColorValues.Text1
														}
													),
													new LatinFont() { Typeface = "+mn-lt" },
													new EastAsianFont() { Typeface = "+mn-ea" },
													new ComplexScriptFont() { Typeface = "+mn-cs" }
												) {
													FontSize = 900,
													Bold = false,
													Italic = false,
													Underline = TextUnderlineValues.None,
													Strike = TextStrikeValues.NoStrike,
													Kerning = 1200,
													Baseline = 0
												}
											),
											new EndParagraphRunProperties() { Language = "en-US" }
										)
									),
									new CrossingAxis() { Val = 48672768U },
									new Crosses() { Val = CrossesValues.AutoZero },
									new AutoLabeled() { Val = true },
									new LabelAlignment() { Val = LabelAlignmentValues.Center },
									new LabelOffset() { Val = 100 },
									new NoMultiLevelLabels() { Val = false }
								)
							);

						// Add the Value Axis.
						var valAx = plotArea
							.AppendChild(
								new ValueAxis(
									new AxisId() { Val = 48672768u },
									new Scaling(new Orientation() {
										Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax
									}),
									new Delete() { Val = false },
									new AxisPosition() { Val = AxisPositionValues.Left },
									new MajorGridlines(
										new ChartShapeProperties(
											new DocumentFormat.OpenXml.Drawing.Outline(
												new SolidFill(
													new SchemeColor(
														new LuminanceModulation() { Val = 15000 },
														new LuminanceOffset() { Val = 85000 }
													) { Val = SchemeColorValues.Text1 }
												),
												new Round()
											) {
												Width = 9525,
												CapType = LineCapValues.Flat,
												CompoundLineType = CompoundLineValues.Single,
												Alignment = PenAlignmentValues.Center
											},
											new EffectList()
										)
									),
									new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() {
										FormatCode = "General",
										SourceLinked = true
									},
									new MajorTickMark() { Val = TickMarkValues.None },
									new MinorTickMark() { Val = TickMarkValues.None },
									new TickLabelPosition() { Val = TickLabelPositionValues.NextTo },
									new ChartShapeProperties(
										new NoFill(),
										new DocumentFormat.OpenXml.Drawing.Outline(new NoFill()),
										new EffectList()
									),
									new DocumentFormat.OpenXml.Drawing.Charts.TextProperties(
										new BodyProperties() {
											Rotation = -60000000,
											UseParagraphSpacing = true,
											VerticalOverflow = TextVerticalOverflowValues.Ellipsis,
											Vertical = TextVerticalValues.Horizontal,
											Wrap = TextWrappingValues.Square,
											Anchor = TextAnchoringTypeValues.Center,
											AnchorCenter = true,
										},
										new Paragraph(
											new ParagraphProperties(
												new DefaultRunProperties(
													new SolidFill(
														new SchemeColor(
															new LuminanceModulation() { Val = 65000 },
															new LuminanceOffset() { Val = 35000 }
														) {
															Val = SchemeColorValues.Text1
														}
													),
													new LatinFont() { Typeface = "+mn-lt" },
													new EastAsianFont() { Typeface = "+mn-ea" },
													new ComplexScriptFont() { Typeface = "+mn-cs" }
												) {
													FontSize = 900,
													Bold = false,
													Italic = false,
													Underline = TextUnderlineValues.None,
													Strike = TextStrikeValues.NoStrike,
													Kerning = 1200,
													Baseline = 0
												}
											),
											new EndParagraphRunProperties() { Language = "en-US" }
										)
									),
									new CrossingAxis() { Val = 48650112U },
									new Crosses() { Val = CrossesValues.AutoZero },
									new CrossBetween() { Val = ch is Models.AreaChart ? CrossBetweenValues.MidpointCategory : CrossBetweenValues.Between })
							);
						// Add the chart Legend.
						//Legend legend = chart.AppendChild(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
						//	new Layout()));

						chart.Append(new PlotVisibleOnly() { Val = true });
						chart.Append(new DisplayBlanksAs() { Val = ch is Models.AreaChart ? DisplayBlanksAsValues.Zero : DisplayBlanksAsValues.Gap });
						chart.Append(new ShowDataLabelsOverMaximum() { Val = false });
						#endregion

						// Save the chart part.
						chartPart.ChartSpace.Save();
					}

					#region position?
					// Position the chart on the worksheet using a TwoCellAnchor object.
					drawingsPart.WorksheetDrawing = new WorksheetDrawing();
					TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());
					twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("0"),
						new ColumnOffset("0"),
						new RowId("0"),
						new RowOffset("0")));
					twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("8"),
						new ColumnOffset("0"),
						new RowId("15"),
						new RowOffset("0")));

					// Append a GraphicFrame to the TwoCellAnchor object.
					var graphicFrame = twoCellAnchor.AppendChild(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
					graphicFrame.Macro = "";

					graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
						new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() {
							Id = new UInt32Value(2u),
							Name = "Chart 1"
						}, new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

					graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
										new Extents() { Cx = 0L, Cy = 0L }));

					graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

					twoCellAnchor.Append(new ClientData());
					#endregion

					// Save the WorksheetDrawing object.
					drawingsPart.WorksheetDrawing.Save();
				}

			}

			var validator = new OpenXmlValidator();
			var errors = validator.Validate(spreadsheetDocument);
			if (errors.Any()) {
				var sbError = new StringBuilder();
				sbError.Append("ERROR: ");
				foreach(var e in errors) {
					sbError.Append($"***{e.Node.ToString()}:{e.Description}***");
				}
				throw new Exception(sbError.ToString());
			}

			workbook.Save();

			// Close the document.
			spreadsheetDocument.Close();
		}

		public MemoryStream Save() {
			var memoryStream = new MemoryStream();
			Save(memoryStream);
			return memoryStream;
		}

		public void Save(Stream stream) {
			Save(SpreadsheetDocument.
					Create(stream, SpreadsheetDocumentType.Workbook));
		}

		public void Save(string path) {
			//Cria um documento para o path (TODO: criar versão stream)
			Save(SpreadsheetDocument.
					Create(path, SpreadsheetDocumentType.Workbook));
		}

		//https://msdn.microsoft.com/en-US/library/office/cc861607.aspx
		// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
		// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
		private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart) {
			// If the part does not contain a SharedStringTable, create one.
			if (shareStringPart.SharedStringTable == null) {
				shareStringPart.SharedStringTable = new SharedStringTable();
			}

			int i = 0;

			// Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
			foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>()) {
				if (item.InnerText == text) {
					return i;
				}

				i++;
			}

			// The text does not exist in the part. Create the SharedStringItem and return its index.
			shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
			shareStringPart.SharedStringTable.Save();
			return i;
		}

		//https://msdn.microsoft.com/en-US/library/office/cc861607.aspx
		// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
		// If the cell already exists, returns it. 
		private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart) {
			Worksheet worksheet = worksheetPart.Worksheet;
			SheetData sheetData = worksheet.GetFirstChild<SheetData>();
			string cellReference = columnName + rowIndex;

			// If the worksheet does not contain a row with the specified row index, insert one.
			Row row;
			if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0) {
				row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
			} else {
				row = new Row() { RowIndex = rowIndex };
				sheetData.Append(row);
			}

			// If there is not a cell with the specified column name, insert one.  
			if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0) {
				return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
			} else {
				// Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
				Cell refCell = null;
				foreach (var cell in row.Elements<Cell>()) {
					if (string.Compare(cell.CellReference.Value, cellReference, true) > 0) {
						refCell = cell;
						break;
					}
				}

				var newCell = new Cell() { CellReference = cellReference };
				row.InsertBefore(newCell, refCell);

				worksheet.Save();
				return newCell;
			}
		}

		static string ColTranslate(int cl) {
			//pega o nome da coluna
			int dividend = cl;
			string columnName = String.Empty;
			int module;

			while (dividend > 0) {
				module = (dividend - 1) % 26;
				columnName = string.Format("{0}{1}", Convert.ToChar(65 + module), columnName);
				dividend = (int)((dividend - module) / 26);
			}

			return columnName;
		}
	}
}
