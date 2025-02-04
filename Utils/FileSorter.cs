using System.Runtime.InteropServices;

public class FileSorter
{
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
    private static extern int StrCmpLogicalW(string x, string y);

    public static int CompareFileNames(string x, string y)
    {
        return StrCmpLogicalW(x, y);
    }
}
