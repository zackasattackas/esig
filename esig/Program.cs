using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace esig
{
    internal class Program
    {
        private static void Main()
        {
            var errorEncountered = false;
            var showingHelp = false;

            Console.WriteLine("\r\nESig - Configure Microsoft Outlook e-mail signatures \r\nZack Bolin (2018)\r\n");

            try
            {
                #region Capture / Validate command line arguments

                var args = ParseCommandLineArguments();

                if (!args.Any() || args.Get("?") != null || args.Get("Help") != null)
                {
                    showingHelp = true;
                    PrintHelp();
                    return;
                }

                var templateFile = args.Get("Template");
                //var variablesFile = args.Get("Variables");
                var newEmailsOnly = args.Get("NewEmails") != null;
                var repliesOnly = args.Get("Replies") != null;
                var officeVersion = args.Get("10") != null ? 14 : args.Get("13") != null ? 15 : args.Get("16") != null ? 16 : (int?) null;
                var persist = args.Get("Persist") != null;

                if (templateFile == null || bool.TryParse(templateFile, out _))
                {
                    throw new ArgumentException("The path to the e-mail signature template file was not provided.");
                }
                if (Path.GetExtension(templateFile).ToLower() != ".docx")
                {
                    throw new ArgumentException("Invalid template file format. The file must be a .DOCX document.");
                }
                if (!File.Exists(templateFile))
                {
                    throw new FileNotFoundException("Failed to locate e-mail signature template file.", templateFile);
                }

                //var variables = new Dictionary<string, string>();

                //if (!string.IsNullOrEmpty(variablesFile))
                //{
                //    if (!File.Exists(variablesFile))
                //    {
                //        throw new FileNotFoundException("Failed to locate variables file.", variablesFile);
                //    }

                //    foreach (var line in File.ReadAllLines(variablesFile))
                //    {
                //        var split = line.Split('=');
                //        variables.Add(split[0], split.Length == 2 ? split[1] : string.Empty);
                //    }
                //}

                #endregion

                var templateCopy = Path.Combine(Path.GetTempPath(), Path.GetFileName(templateFile));

                File.Copy(templateFile, templateCopy, true);

                var winWord = new Application();
                winWord.Documents.Open(templateCopy);

                var user = ADUser.GetCurrentUser();

                //var displayName = variables.TryGetValue("DisplayName", out var name) ? name : user?.DisplayName;
                //var email = variables.TryGetValue("DisplayName", out var mail) ? mail : user?.EmailAddress;
                //var title = variables.TryGetValue("DisplayName", out var desc) ? desc : user?.Title;

                var displayName = user?.DisplayName;
                var email = user?.EmailAddress;
                var title = user?.Title;

                FindAndReplace(ref winWord, "{DisplayName}", displayName);
                FindAndReplace(ref winWord, "{Email}", email ?? "help@krns-inc.com");   // Defaults to our company's support email address
                FindAndReplace(ref winWord, "{Title}", title ?? "Junior Employee");     // 'The Office' reference, anyone?

                var signaturesPath = GetSignaturesDirectory();
                if (!signaturesPath.Exists)
                {
                    if (signaturesPath.Parent != null && signaturesPath.Parent.Exists)
                    {
                        signaturesPath.Parent.CreateSubdirectory(signaturesPath.Name);
                    }
                    else
                    {
                        throw new DirectoryNotFoundException($"One or more directories in the path \"{signaturesPath.FullName}\" was not found. This could mean Microsoft Office is not installed or has not run completed first-run tasks.");
                    }
                }
                SignatureTypes signatureType;
                if (newEmailsOnly)
                    signatureType = SignatureTypes.NewSignature;
                else if (repliesOnly)
                    signatureType = SignatureTypes.ReplySignature;
                else
                    signatureType = SignatureTypes.NewSignature | SignatureTypes.ReplySignature;

                var fileName = GetSignatureFileNameWithoutExtension(signatureType);

                SaveAs(ref winWord, WdSaveFormat.wdFormatHTML, Path.Combine(signaturesPath.FullName, $"{fileName}.htm"));
                SaveAs(ref winWord, WdSaveFormat.wdFormatRTF, Path.Combine(signaturesPath.FullName, $"{fileName}.rtf"));
                SaveAs(ref winWord, WdSaveFormat.wdFormatText, Path.Combine(signaturesPath.FullName, $"{fileName}.txt"));

                var emailOptions = winWord.EmailOptions;
                if (newEmailsOnly)
                {
                    emailOptions.EmailSignature.NewMessageSignature = fileName;
                }
                else if (repliesOnly)
                {
                    emailOptions.EmailSignature.ReplyMessageSignature = fileName;
                }
                else
                {
                    emailOptions.EmailSignature.NewMessageSignature = fileName;
                    emailOptions.EmailSignature.ReplyMessageSignature = fileName;
                }

                if (persist)
                {
                    if (officeVersion == null)
                    {
                        throw new ArgumentException("No Microsoft Office version was specified.");
                    }
                    SetSignatureInRegistry(fileName, signatureType, officeVersion.Value);
                }

                winWord.ActiveDocument.Close();
                winWord.Quit();

                Marshal.ReleaseComObject(emailOptions);
                Marshal.ReleaseComObject(winWord);

                File.Delete(templateCopy);
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("ERROR: One or more arguments were typed incorrectly. See -? for usage syntax.");
            }
            catch (Exception e)
            {
                errorEncountered = true;
                Console.WriteLine($"ERROR: {e.Message}");
            }
            finally
            {
                if (!showingHelp)
                {
                    Console.WriteLine($"The command completed {(errorEncountered ? "with errors" : "successfully")}.");
                }
#if DEBUG
                Console.Write("\r\nPress any key to exit...");
                Console.ReadKey();
#endif
            }
        }

        public static Dictionary<string,string> ParseCommandLineArguments()
        {
            var args = Environment.GetCommandLineArgs();
            var argsDictionary = new Dictionary<string,string>(StringComparer.CurrentCultureIgnoreCase);
            for (var i = 1; i < args.Length; i++)
            {
                if (char.IsLetterOrDigit(args[i][0]))
                {
                    continue;
                }
                var arg = args[i].Substring(1);
                var value = args.Length > i + 1 && char.IsLetterOrDigit(args[i + 1][0]) ? args[++i] : bool.TrueString;
                argsDictionary.Add(arg,  value);
            }

            return argsDictionary;
        }
        public static void PrintHelp()
        {
            const string fmt = "    {0,-10} {1}";
            //Console.WriteLine("USAGE: esig -Template <filePath> [-Variables <filePath>] [-NewEmails | -Replies] -10 | -13 | -16\r\n");
            Console.WriteLine("USAGE: esig -Template <filePath> [-NewEmails | -Replies] [-10 | -13 | -16 [-Persist]]\r\n");
            Console.WriteLine(GetWrappedText(10, "    Note: Square brackets indicate optional parameters. Parameter names are not case sensitive, and can be specified in any order.", false));
            Console.WriteLine(fmt, "-Template", "The path to the template file.\r\n");
            //Console.WriteLine(fmt, "-Variables", GetWrappedText(15, "The path to a text file containing \"key=value\" pairs. These values will override those pulled from the ADSI.", false));
            Console.WriteLine(fmt, "-NewEmails", "The signature should be applied to new emails only.\r\n");
            Console.WriteLine(fmt, "-Replies", "The signature should be applied to replies only.\r\n");
            Console.WriteLine(fmt, "-<version>", "The version of Microsoft Office installed. -10 = 2010; -13 = 2013; -16 = 2016\r\n");
            Console.WriteLine(fmt, "-Persist", GetWrappedText(15, "Configures a registry value which prevents the user from changing their signature manually in Outlook.", false));            
        }
        public static void FindAndReplace(ref Application word, string searchPattern, string replacement)
        {
            word.Selection.Find.Execute(searchPattern, true, true, false, false, false, true, 1, false, replacement, 2);
        }
        public static DirectoryInfo GetSignaturesDirectory()
        {
            return new DirectoryInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft\\Signatures"));
        }

        public static string GetSignatureFileNameWithoutExtension(SignatureTypes signatureType)
        {
            return $"{Environment.UserName} - {(signatureType == SignatureTypes.NewSignature ? "New Message" : signatureType == SignatureTypes.ReplySignature ? "Replies" : "All Mail")}";
        }
        public static void SaveAs(ref Application word, WdSaveFormat format, string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            word.ActiveDocument.SaveAs(fileName, format);            
        }

        public static void SetSignatureInRegistry(string signatureName, SignatureTypes signatureType, int officeVersion)
        {
            var keyName = $"Software\\Microsoft\\Office\\{officeVersion}.0\\Common\\MailSettings";
            var registryKey = Registry.CurrentUser.OpenSubKey(keyName, true) ?? throw new Exception($"The key HKCU:\\{keyName} was not found.");
            if (signatureType == (SignatureTypes.NewSignature | SignatureTypes.ReplySignature))
            {
                registryKey.SetValue(nameof(SignatureTypes.NewSignature), signatureName, RegistryValueKind.String);
                registryKey.SetValue(nameof(SignatureTypes.ReplySignature), signatureName, RegistryValueKind.String);
            }
            registryKey.SetValue(signatureType.ToString(), signatureName, RegistryValueKind.String);
        }
        public static string GetWrappedText(int leftPad, string input, bool padFirstLine = true)
        {
            return GetWrappedTextLocal(input, true);

            string GetWrappedTextLocal(string text, bool firstCall = false)
            {
                var maxWidth = Console.BufferWidth - leftPad;
                var padStr = new string(' ', leftPad);

                if (text.Length < maxWidth)
                {
                    return firstCall ? $"{(padFirstLine ? padStr : string.Empty)}{text}" : padStr + text;
                }

                var sb = new StringBuilder();

                var fittedLine = text.Substring(0, text.Substring(0, maxWidth).LastIndexOf(" ", StringComparison.Ordinal));

                sb.AppendLine(firstCall ? $"{(padFirstLine ? padStr : string.Empty)}{fittedLine}" : padStr + fittedLine);

                return (text.Length > fittedLine.Length
                    ? sb.AppendLine(GetWrappedTextLocal(text.Substring(fittedLine.Length + 1)))
                    : sb).ToString();
            }
        }
    }
    [Flags]
    public enum SignatureTypes
    {
        NewSignature,
        ReplySignature
    }
    // ReSharper disable once InconsistentNaming
    public class ADUser
    {
        public string DisplayName { get; set; }
        public string EmailAddress { get; set; }
        public string Title { get; set; }

        private ADUser(string name, string email, string title)
        {
            DisplayName = name;
            EmailAddress = email;
            Title = title;
        }

        public static ADUser GetCurrentUser()
        {
            var filter = $"(&(objectCategory=User)(samAccountName={Environment.UserName}))";
            var searcher = new DirectorySearcher { Filter = filter };
            var user = searcher.FindOne();

            try
            {
                return new ADUser(
                    user.Properties["displayname"][0].ToString(),
                    user.Properties["mail"][0].ToString(),
                    user.Properties["description"][0].ToString());
            }
            catch
            {
                throw new Exception("Failed to retrieve one or more Active Directory attributes for the current user.");
            }
        }
    }

    public static class Extensions
    {
        public static string Get(this Dictionary<string,string> d, string name)
        {
            return d.TryGetValue(name, out var value) ? value : null;
        }
    }
}
