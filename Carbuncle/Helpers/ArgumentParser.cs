using System.Collections.Generic;
using System.Diagnostics;
using System;

namespace Carbuncle
{
    //Argument parser based on the code created by @harmj0y as part of Rubeus
    //https://github.com/GhostPack/Rubeus/blob/master/Rubeus/Domain/ArgumentParser.cs
    public static class ArgumentParser
    {
        public static ArgumentParserResult Parse(IEnumerable<string> args)
        {
            var arguments = new Dictionary<string, string>();
            try
            {
                foreach (var argument in args)
                {
                    var idx = argument.IndexOf(':');
                    if (idx > 0)
                        arguments[argument.Substring(0, idx).Replace("/","")] = argument.Substring(idx + 1);
                    else
                        arguments[argument.Replace("/","")] = string.Empty;
                }

                return ArgumentParserResult.Success(arguments);
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return ArgumentParserResult.Failure();
            }
        }
    }
    public class ArgumentParserResult
    {
        public bool ParsedOk { get; }
        public Dictionary<string, string> Arguments { get; }

        private ArgumentParserResult(bool parsedOk, Dictionary<string, string> arguments)
        {
            ParsedOk = parsedOk;
            Arguments = arguments;
        }

        public static ArgumentParserResult Success(Dictionary<string, string> arguments)
            => new ArgumentParserResult(true, arguments);

        public static ArgumentParserResult Failure()
            => new ArgumentParserResult(false, null);

    }
}