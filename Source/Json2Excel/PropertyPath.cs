using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace Json2Excel
{
	class PropertyPath : IComparable<PropertyPath>, IEquatable<PropertyPath>
	{
		private string PropertyName { get; set; }
		private PropertyPath ParentProperty { get; set; }

		public PropertyPath(string propertyName)
		{
			PropertyName = propertyName;
		}

		public PropertyPath(PropertyPath parent, string propertyName)
			: this(propertyName)
		{
			ParentProperty = parent;
		}

		public JToken GetValue(JObject obj)
		{
			if (ParentProperty != null)
			{
				var jv = ParentProperty.GetValue(obj);

				if (jv?.Type == JTokenType.Object)
				{
					var obj2 = (JObject)jv;
					return obj2.GetValue(PropertyName);
				}

				return null;
			}

			return obj.GetValue(PropertyName);
		}

		public int CompareTo(PropertyPath other)
		{
			if (ParentProperty == null)
			{
				if (other.ParentProperty != null) return 1;
			}
			else
			{
				if (other.ParentProperty == null) return -1;
				var result = ParentProperty.CompareTo(other);
				if (result != 0) return result;
			}

			return string.Compare(PropertyName, other.PropertyName, StringComparison.Ordinal);
		}

		public override string ToString()
		{
			if (ParentProperty != null) return ParentProperty + "." + PropertyName;
			return PropertyName;
		}

		public static List<PropertyPath> GetPropertyPaths(JObject obj)
		{
			var list = new List<PropertyPath>();
			GetPropertyPaths(list, obj, null);
			return list;
		}

		private static void GetPropertyPaths(List<PropertyPath> list, JObject obj, PropertyPath parent)
		{
			foreach (JProperty jProperty in obj.Properties())
			{
				switch (jProperty.Value.Type)
				{
					case JTokenType.Object:
						GetPropertyPaths(list, (JObject)jProperty.Value, new PropertyPath(parent, jProperty.Name));
						break;
					case JTokenType.Array:
						break;
					default:
						list.Add(new PropertyPath(parent, jProperty.Name));
						break;
				}
			}
		}

		public bool Equals(PropertyPath other)
		{
			if (ReferenceEquals(null, other)) return false;
			if (ReferenceEquals(this, other)) return true;
			return string.Equals(PropertyName, other.PropertyName) && Equals(ParentProperty, other.ParentProperty);
		}

		public override bool Equals(object obj)
		{
			if (ReferenceEquals(null, obj)) return false;
			if (ReferenceEquals(this, obj)) return true;
			if (obj.GetType() != this.GetType()) return false;
			return Equals((PropertyPath)obj);
		}

		public override int GetHashCode()
		{
			unchecked
			{
				return ((PropertyName != null ? PropertyName.GetHashCode() : 0) * 397) ^ (ParentProperty != null ? ParentProperty.GetHashCode() : 0);
			}
		}
	}
}