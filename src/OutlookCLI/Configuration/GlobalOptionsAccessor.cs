namespace OutlookCLI.Configuration;

public static class GlobalOptionsAccessor
{
    private static readonly AsyncLocal<GlobalOptions> _current = new();

    public static GlobalOptions Current
    {
        get => _current.Value ?? new GlobalOptions();
        set => _current.Value = value;
    }
}
