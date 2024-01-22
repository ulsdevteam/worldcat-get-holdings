using CommandLine;

class Options
{
    [Value(0)]
    public string InputFile { get; set; }

    [Option('f', "from")]
    public int? From { get; set; }

    [Option('t', "to")]
    public int? To { get; set; }
}