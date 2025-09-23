# --------------------------------------------------------------
# 1. Assemblies we need
# --------------------------------------------------------------
$assemblies = @(
    'System.Windows.Forms',   # ListView, SortOrder, etc.
    'System.Collections'      # IComparer interface
)

# --------------------------------------------------------------
# 2. Declare the C# class
# --------------------------------------------------------------
Add-Type -ReferencedAssemblies $assemblies `
         -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Collections;

public class ListViewColumnSorter : IComparer
{
    private int   _columnToSort   = 0;
    private SortOrder _order       = SortOrder.None;
    private readonly StringComparer _comparer =
        StringComparer.CurrentCultureIgnoreCase;   // case‑insensitive

    public int Compare(object x, object y)
    {
        var itemX = x as ListViewItem;
        var itemY = y as ListViewItem;
        if (itemX == null || itemY == null) return 0;

        string s1 = itemX.SubItems.Count > _columnToSort ?
                     itemX.SubItems[_columnToSort].Text : string.Empty;
        string s2 = itemY.SubItems.Count > _columnToSort ?
                     itemY.SubItems[_columnToSort].Text : string.Empty;

        int result = _comparer.Compare(s1, s2);
        return _order switch
        {
            SortOrder.Ascending  =>  result,
            SortOrder.Descending => -result,
            _                    =>  0
        };
    }

    public int SortColumn
    {
        get => _columnToSort;
        set => _columnToSort = value;
    }

    public SortOrder Order
    {
        get => _order;
        set => _order = value;
    }
}
"@

# --------------------------------------------------------------
# 3.  Instantiate the sorter (works in all hosts)
# --------------------------------------------------------------
#$sorter = [ListViewColumnSorter]::new()

# ------------------------------------------------------------------
# Example: attach the sorter to a ListView instance
# ------------------------------------------------------------------
# $listView.ListViewItemSorter = $sorter
# $listView.Sort()