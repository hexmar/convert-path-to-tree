using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WA_Test_V5.Models
{
	public class ItemWithCid : IEquatable<ItemWithCid>
	{
		public int Cid;
		public string Name;

		public ItemWithCid(int cid, string name)
		{
			Cid = cid;
			Name = name;
		}

		public bool Equals(ItemWithCid other)
		{
			return Cid.Equals(other.Cid) && Name.Equals(other.Name);
		}

		public override int GetHashCode()
		{
			return Cid.GetHashCode() ^ Name.GetHashCode();
		}
	}
}