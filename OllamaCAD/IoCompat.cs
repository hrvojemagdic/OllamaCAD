using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace OllamaCAD
{   
    /// <summary>
    /// Provides async-compatible file I/O helpers for .NET Framework.
    ///
    /// - Wraps synchronous System.IO file operations inside Task.Run.
    /// - Enables non-blocking usage from async workflows in the add-in.
    /// - Used for reading/writing memory files (chat turns, summaries, settings).
    ///
    /// This ensures UI responsiveness when performing disk operations.
    /// </summary>
    internal static class IoCompat
    {
        public static Task<string> ReadAllTextAsync(string path, Encoding enc)
        {
            return Task.Run(() => File.ReadAllText(path, enc));
        }

        public static Task WriteAllTextAsync(string path, string text, Encoding enc)
        {
            return Task.Run(() => File.WriteAllText(path, text, enc));
        }

        public static Task AppendAllTextAsync(string path, string text, Encoding enc)
        {
            return Task.Run(() => File.AppendAllText(path, text, enc));
        }

        public static Task<string[]> ReadAllLinesAsync(string path, Encoding enc)
        {
            return Task.Run(() => File.ReadAllLines(path, enc));
        }
    }
}
