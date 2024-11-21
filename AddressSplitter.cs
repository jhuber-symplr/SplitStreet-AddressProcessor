using System.Text.RegularExpressions;

public static class AddressSplitter
{
    public static (string Address1, string Address2) SplitAddress(string address, string pattern)
    {
        address = Regex.Replace(address, @"\s{2,}", " ");

        int commaIndex = address.IndexOf(',');
        if (commaIndex >= 0)
        {
            return (address.Substring(0, commaIndex).Trim(), address.Substring(commaIndex + 1).Trim());
        }

        var match = Regex.Match(address, pattern, RegexOptions.IgnoreCase);
        if (match.Success)
        {
            int index = match.Index;
            return (address.Substring(0, index).Trim(), address.Substring(index).Trim());
        }

        return (address, "");
    }
}
