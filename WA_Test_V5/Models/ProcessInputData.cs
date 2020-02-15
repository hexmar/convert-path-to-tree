using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using WA_Test_V5.Interface.TreeView;

namespace WA_Test_V5.Models
{
	public class ProcessInputData
	{
		private ExcelPackage pack;

		public ProcessInputData(string filePath)
		{
			FileInfo fInfo = new FileInfo(filePath);
			if (fInfo.Exists != true) throw new Exception();
			pack = new ExcelPackage(fInfo);
		}

		public List<TreeViewElements> GetData()
		{
			var sheets = pack.Workbook.Worksheets;
			var dataSheet = sheets.First();
			var numberOfRows = dataSheet.Dimension.End.Row;
			var numberOfCols = dataSheet.Dimension.End.Column;
			ExcelRange Cells = dataSheet.Cells;

			var rootNode = new Dictionary<string, object>();
			for (var row = 2; row < numberOfRows; row++)
			{
				var node = rootNode;
				for (var col = 1; col < numberOfCols - 2; col++)
				{
					var cellValue = Cells[row, col].Value.ToString();
					object nextNode;
					var nodeExists = node.TryGetValue(cellValue, out nextNode);
					if (!nodeExists)
					{
						nextNode = new Dictionary<string, object>();
						node.Add(cellValue, nextNode);
					}

					node = (Dictionary<string, object>)nextNode;
				}

				var lastCoordinate = Cells[row, numberOfCols - 2].Value.ToString();
				object existingSet;
				var setExists = node.TryGetValue(lastCoordinate, out existingSet);
				if (!setExists)
				{
					existingSet = new HashSet<ItemWithCid>(EqualityComparer<ItemWithCid>.Default);
					node.Add(lastCoordinate, existingSet);
				}

				var set = (HashSet<ItemWithCid>)existingSet;

				var cidItem = new ItemWithCid(
					Cells[row, numberOfCols].Value.ToString(),
					Cells[row, numberOfCols - 1].Value.ToString());
				set.Add(cidItem);
			}

			return new List<TreeViewElements>();
		}
	}
}