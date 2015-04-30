using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;

using IWshRuntimeLibrary;
using System.Runtime.InteropServices;

namespace ModEnv
{
	public enum Options
	{
		Add, AddEnd, Remove, RemoveAny
	}

	class ModifyEnviron
	{
		const string PathVar = "Path";

		private Options _myOpts;
		private bool _isDebug = false;
		private bool _registryChange = false;
		private string _batchFileName = null;
		private List<string> _addlEntries = new List<string>();
		private List<string> _paths;

		public ModifyEnviron(string[] args)
		{
			ParseArgs(args);
		}

		private void ParseArgs(string[] args)
		{
			_myOpts = Options.Add;

			int argc = 0;
			while (argc < args.Length)
			{
				string p = args[argc++];
				if (p[0] == '-' || p[0] == '/')
				{
					string opt = p.Substring(1).ToLowerInvariant();
					switch (opt)
					{
						case "add":
							_myOpts = Options.Add;
							break;

						case "addend":
							_myOpts = Options.AddEnd;
							break;

						case "makebat":
							if (argc < args.Length)
							{
								_batchFileName = Environment.ExpandEnvironmentVariables(args[argc++]);
							}
							break;

						case "registry":
							_registryChange = true; break;

						case "debug":
							_isDebug = true; break;

						case "remove":
							_myOpts = Options.Remove;
							break;

						case "removeany":
							_myOpts = Options.RemoveAny;
							break;

						default:
							break;
					}
				}
				else
				{
					_addlEntries.Add(Environment.ExpandEnvironmentVariables(p)); // get environment expanded
				}
			}
		}

		private void ReadPath()
		{
			string curPath = Environment.GetEnvironmentVariable(PathVar);
			if (_isDebug)
				Console.WriteLine("Original Path: '{0}'\n", curPath);

			string[] curPaths = curPath.Split(new char[] { ';' });
			_paths = curPaths.ToList();
		}

		private void RemoveFromPaths(List<string> removeItems, bool inexact)
		{
			removeItems.ForEach(s =>
			{
				// remove all strings matching beginning of item
				_paths.RemoveAll(delegate(string match)
				{
					return (!inexact) ?
						string.Equals(s, match, StringComparison.CurrentCultureIgnoreCase) :
						(match.IndexOf(s, StringComparison.CurrentCultureIgnoreCase) >= 0);
				});
			});
		}

		public void Operate()
		{
			int addlCount = _addlEntries.Count;
			if (addlCount == 0)
				return;

			// First read all entries in
			ReadPath();

			// Remove entries from path
			RemoveFromPaths(_addlEntries, (_myOpts == Options.RemoveAny));

			if (_myOpts == Options.Add)
			{
				// Add entries to front
				for (int i = addlCount - 1; i >= 0; i--)
				{
					string item = _addlEntries[i];
					_paths.Insert(0, item);
				}
			}
			else if (_myOpts == Options.AddEnd)
			{
				// Add entries to back
				_addlEntries.ForEach(s => _paths.Add(s));
			}
			else if (_myOpts == Options.Remove)
			{
				// remove entries (do nothing)
			}

			// Combine string back
			string finalPath = String.Join(Path.PathSeparator.ToString(), _paths);
			if (_isDebug)
				Console.WriteLine("New Path: '{0}'\n", finalPath);

			// Only write to registry if selected
			if (_registryChange)
			{
				// Now write it to System registry
				SetSystemVariable(PathVar, finalPath, false);
				BroadcastChange();
			}

			if (_batchFileName != null)
			{
				string pathOut = string.Format("@echo off\nSet {0}={1}\n\n", PathVar, finalPath);
				System.IO.File.WriteAllText(_batchFileName, pathOut);
			}
		}

		#region Windows WSH internal functions

		// Functions taken from Greg Houston's blog:
		// http://ghouston.blogspot.com/2005/08/how-to-create-and-change-environment.html

		public static void SetUserVariable(string name, string value, bool isRegExpandSz)
		{
			SetVariable("HKEY_CURRENT_USER\\Environment\\" + name, value, isRegExpandSz);
		}

		public static void SetSystemVariable(string name, string value, bool isRegExpandSz)
		{
			SetVariable("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment\\" + name, value, isRegExpandSz);
		}

		private static void SetVariable(string fullpath, string value, bool isRegExpandSz)
		{
			object objValue = value;
			object objType = (isRegExpandSz) ? "REG_EXPAND_SZ" : "REG_SZ";
			WshShell shell = new WshShell();
			shell.RegWrite(fullpath, ref objValue, ref objType);

			// need to call BroadcastChange() to modify registry
		}

		private static void BroadcastChange()
		{
			int result;

			string envStr = "Environment";
			IntPtr ptrA = Marshal.StringToHGlobalAnsi(envStr);
			IntPtr ptrU = Marshal.StringToHGlobalUni(envStr);

			SendMessageTimeout((System.IntPtr)HWND_BROADCAST,
				WM_SETTINGCHANGE, IntPtr.Zero, ptrA, SMTO_ABORTIFHUNG 
				, 5000, out result);
			SendMessageTimeout((System.IntPtr)HWND_BROADCAST,
				WM_SETTINGCHANGE, IntPtr.Zero, ptrU, SMTO_ABORTIFHUNG 
				, 5000, out result);

			Marshal.FreeHGlobal(ptrA);
			Marshal.FreeHGlobal(ptrU);
		}

		[DllImport("user32.dll",
			 CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		public static extern bool
			SendMessageTimeout(
			IntPtr hWnd,
			int Msg,
			IntPtr wParam,
			IntPtr lParam,
			int fuFlags,
			int uTimeout,
			out int lpdwResult
			);

		public const int HWND_BROADCAST = 0xffff;
		public const int WM_SETTINGCHANGE = 0x001A;
		public const int SMTO_NORMAL = 0x0000;
		public const int SMTO_BLOCK = 0x0001;
		public const int SMTO_ABORTIFHUNG = 0x0002;
		public const int SMTO_NOTIMEOUTIFNOTHUNG = 0x0008;
		#endregion

		static void Usage()
		{
			Console.WriteLine("ModEnv [/option] [path1] .. [pathn]: Modify Environment");
			Console.WriteLine("Options:\n/Add - add to front of PATH(default)\n/AddEnd - add to rear of PATH");
			Console.WriteLine("/MakeBat - [output_file] (create bat file)");
			Console.WriteLine("/Registry - write to System registry");
			Console.WriteLine("/Remove - remove from PATH\n/RemoveAny - remove any matching string from PATH\n");
			System.Environment.Exit(0);
		}

		static void Main(string[] args)
		{
			if (args.Length == 0)
				Usage();

			ModifyEnviron work = new ModifyEnviron(args);
			work.Operate();
		}
	}
}
