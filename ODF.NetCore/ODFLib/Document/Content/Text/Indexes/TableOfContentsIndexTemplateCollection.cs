/*
 * $Id: TableOfContentsIndexTemplateCollection.cs,v 1.1 2006/01/29 11:28:22 larsbm Exp $
 */

/*
 * License: 
 * GNU Lesser General Public License. You should recieve a
 * copy of this within the library. If not you will find
 * a whole copy at http://www.gnu.org/licenses/lgpl.html .
 * 
 * Author:
 * Copyright 2006, Lars Behrmann, lb@OpenDocument4all.com
 * 
 * Last changes:
 * 
 */

using System.Collections;
using ODFLib.Document.Collections;

namespace ODFLib.Document.Content.Text.Indexes
{
	/// <summary>
	/// TableOfContentsIndexTemplateCollection
	/// </summary>
	public class TableOfContentsIndexTemplateCollection : CollectionWithEvents
	{
		/// <summary>
		/// Adds the specified value.
		/// </summary>
		/// <param name="value">The value.</param>
		/// <returns></returns>
		public int Add(ODFLib.Document.Content.Text.Indexes.TableOfContentsIndexTemplate value)
		{
			return base.List.Add(value as object);
		}

		/// <summary>
		/// Removes the specified value.
		/// </summary>
		/// <param name="value">The value.</param>
		public void Remove(ODFLib.Document.Content.Text.Indexes.TableOfContentsIndexTemplate value)
		{
			base.List.Remove(value as object);
		}

		/// <summary>
		/// Inserts the specified index.
		/// </summary>
		/// <param name="index">The index.</param>
		/// <param name="value">The value.</param>
		public void Insert(int index, ODFLib.Document.Content.Text.Indexes.TableOfContentsIndexTemplate value)
		{
			base.List.Insert(index, value as object);
		}

		/// <summary>
		/// Determines whether [contains] [the specified value].
		/// </summary>
		/// <param name="value">The value.</param>
		/// <returns>
		/// 	<c>true</c> if [contains] [the specified value]; otherwise, <c>false</c>.
		/// </returns>
		public bool Contains(ODFLib.Document.Content.Text.Indexes.TableOfContentsIndexTemplate value)
		{
			return base.List.Contains(value as object);
		}

		/// <summary>
		/// Gets the <see cref="TableOfContentsIndexTemplate"/> at the specified index.
		/// </summary>
		/// <value></value>
		public ODFLib.Document.Content.Text.Indexes.TableOfContentsIndexTemplate this[int index]
		{
			get { return (base.List[index] as ODFLib.Document.Content.Text.Indexes.TableOfContentsIndexTemplate); }
		}
	}
}

/*
 * $Log: TableOfContentsIndexTemplateCollection.cs,v $
 * Revision 1.1  2006/01/29 11:28:22  larsbm
 * - Changes for the new version. 1.2. see next changelog for details
 *
 * Revision 1.1  2006/01/05 10:31:10  larsbm
 * - AODL merged cells
 * - AODL toc
 * - AODC batch mode, splash screen
 *
 */