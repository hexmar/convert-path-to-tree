using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using WA_Test_V5.Interface.TreeView;

namespace WA_Test_V5.Models
{
	public class ProcessInputData : IDisposable
	{
		private ExcelPackage pack;
		private int numberOfRows;
		private int numberOfCols;
		private ExcelRange cells;
		private int idCounter;

		public ProcessInputData(string filePath)
		{
			FileInfo fInfo = new FileInfo(filePath);
			if (fInfo.Exists != true) throw new Exception();
			pack = new ExcelPackage(fInfo);

			var sheets = pack.Workbook.Worksheets;
			var dataSheet = sheets.First();
			numberOfRows = dataSheet.Dimension.End.Row;
			numberOfCols = dataSheet.Dimension.End.Column;
			cells = dataSheet.Cells;
		}

		public void Dispose()
		{
			pack.Dispose();
		}

		public List<TreeViewElements> GetData()
		{
			var rootNode = GetTree();

			var nodes = GetList(rootNode);

			return nodes;
		}

		private Dictionary<string, object> GetTree()
		{
			var rootNode = new Dictionary<string, object>();
			for (var row = 2; row <= numberOfRows; row++)
			{
				var node = GetNodeLocation(rootNode, row);

				var set = GetLeaveSet(node, row);

				AddNode(set, row);
			}

			return rootNode;
		}

		private Dictionary<string, object> GetNodeLocation(Dictionary<string, object> rootNode, int row)
		{
			var node = rootNode;
			for (var col = 1; col < numberOfCols - 2; col++)
			{
				var cellValue = cells[row, col].Value.ToString();
				object nextNode;
				var nodeExists = node.TryGetValue(cellValue, out nextNode);
				if (!nodeExists)
				{
					nextNode = new Dictionary<string, object>();
					node.Add(cellValue, nextNode);
				}

				node = (Dictionary<string, object>)nextNode;
			}

			return node;
		}

		private HashSet<ItemWithCid> GetLeaveSet(Dictionary<string, object> node, int row)
		{
			var lastCoordinate = cells[row, numberOfCols - 2].Value.ToString();
			object existingSet;
			var setExists = node.TryGetValue(lastCoordinate, out existingSet);
			if (!setExists)
			{
				existingSet = new HashSet<ItemWithCid>(EqualityComparer<ItemWithCid>.Default);
				node.Add(lastCoordinate, existingSet);
			}

			return (HashSet<ItemWithCid>)existingSet;
		}

		private void AddNode(HashSet<ItemWithCid> set, int row)
		{
			int cid;
			var validCid = int.TryParse(cells[row, numberOfCols].Value.ToString(), out cid);
			if (validCid)
			{
				var cidItem = new ItemWithCid(cid, cells[row, numberOfCols - 1].Value.ToString());
				set.Add(cidItem);
			}
		}

		private List<TreeViewElements> GetList(Dictionary<string, object> rootNode)
		{
			var nodes = new List<TreeViewElements>();
			var leaves = new List<TreeViewElements>();
			idCounter = 0;

			var queue = new Queue<KeyValuePair<string, Dictionary<string, object>>>();
			queue.Enqueue(new KeyValuePair<string, Dictionary<string, object>>(
				(idCounter++).ToString(),
				rootNode));

			while (queue.Count > 0)
			{
				var pair = queue.Dequeue();

				var keys = pair.Value.Keys.OrderBy(key => key);
				foreach (var key in keys)
				{
					var currentTreeElement = CreateTreeElement(key, pair.Key);
					nodes.Add(currentTreeElement);

					var nextNode = pair.Value[key];
					if (nextNode is HashSet<ItemWithCid>)
					{
						var set = (HashSet<ItemWithCid>)nextNode;
						AddLeavesToList(set, leaves, currentTreeElement.ID);

						continue;
					}

					var dictionary = (Dictionary<string, object>)nextNode;
					queue.Enqueue(new KeyValuePair<string, Dictionary<string, object>>(
						currentTreeElement.ID,
						dictionary));
				}
			}

			foreach (var element in leaves)
			{
				element.ID = (idCounter++).ToString();
				nodes.Add(element);
			}

			return nodes;
		}

		private TreeViewElements CreateTreeElement(string name, string parentId)
		{
			var currentTreeElement = new TreeViewElements
			{
				ID = (idCounter++).ToString(),
				Name = name,
				CID = -2,
				Parent_ID = parentId,
			};

			return currentTreeElement;
		}

		private void AddLeavesToList(HashSet<ItemWithCid> leaves, List<TreeViewElements> list, string parentId)
		{
			var treeElements = leaves.Select(item => new TreeViewElements
			{
				Parent_ID = parentId,
				Name = item.Name,
				CID = item.Cid
			});
			foreach (var element in treeElements)
			{
				var insertIndex = list.FindIndex(addedElement => string.Compare(addedElement.Name, element.Name) > 0);
				if (insertIndex < 0)
				{
					list.Add(element);
				}
				else
				{
					list.Insert(insertIndex, element);
				}
			}
		}
	}
}