using MimeKit;
using System.IO;
using System.Net.Mail;
using System.Text;

#region Prompts & text

string _promptSourceFolder = "Enter source folder:";
string _promptSourceFolderNotFound = "Source folder not found, enter valid source folder:";
string _promptDestinationFolder = "Enter destination folder:";
string _promptDestionationFolderNotFound = "Destination folder not found, enter valid source folder:";
string _titleStartExtentions = "Processing files to set extensions";
string _promptExtension = "Enter extension for files:";
string _promptExtensionIsEmpty = "Please give an extension";

string getSourceFolder()
{
	Console.WriteLine();
	Console.WriteLine(_promptSourceFolder);
	string? _target = Console.ReadLine();
	while (!System.IO.Directory.Exists(_target))
	{
		Console.WriteLine(_promptSourceFolderNotFound);
		_target = getSourceFolder();
	}
	return _target;
}
string getDestinationFolder()
{
	Console.WriteLine();
	Console.Write(_promptDestinationFolder);
	string? _destination = Console.ReadLine();
	while (!System.IO.Directory.Exists(_destination))
	{
		Console.WriteLine(_promptDestionationFolderNotFound);
		_destination = getDestinationFolder();
	}
	return _destination;
}
string getExtension()
{
	Console.WriteLine();
	Console.WriteLine(_promptExtension);
	string? _extension = Console.ReadLine();
	while (_extension == null || _extension?.Length <= 0)
	{
		Console.WriteLine(_promptExtensionIsEmpty);
		_extension = getExtension();
	}
	return _extension ?? "";
}
#endregion

#region Menu
async void showMenu()
{
	Console.WriteLine("[E]mail");
	Console.WriteLine("m[B]oxes");
	Console.WriteLine("[M]ove");
	Console.WriteLine("Exten[S]ion[S]");
	Console.WriteLine("e[X]it");
	ConsoleKeyInfo _input = Console.ReadKey();
	switch (_input.Key.ToString().ToUpper())
	{
		case "E": // E:\Email\Test
			StartEmail(true);
			showMenu();
			break;
		case "M":
			StartMove();
			showMenu();
			break;
		case "S":
			StartExtensions();
			showMenu();
			break;
		case "B":
			StartMboxes();
			showMenu();
			break;
		case "X":
			Environment.Exit(0);
			break;
		default:
			showMenu();
			break;
	}
}
#endregion


#region Extensions
async Task StartExtensions()
{
	Console.WriteLine(_titleStartExtentions);
	string _extension = getExtension();
	string _target = getSourceFolder();
	string _destination = getDestinationFolder();
	Console.WriteLine($"Going from {_target} to {_destination}");
	await processDirectoryExtensions(_target, _destination, _extension);
}

async Task processDirectoryExtensions(string Path, String Destination, String Extension)
{
	List<Task> TaskList = new List<Task>();
	Console.WriteLine($"Processing directory {Path}");
	if (System.IO.Directory.Exists(Path))
	{
		foreach (string f in System.IO.Directory.GetFiles(Path))
		{
			var LastTask = Task.Run(async () => await ProcessExtension(f, Destination, Extension));
			TaskList.Add(LastTask);
		}
		if (System.IO.Directory.GetDirectories(Path).Length > 0)
		{
			foreach (var v in System.IO.Directory.GetDirectories(Path))
			{
				await processDirectoryExtensions(v, Destination, Extension);
			}
		}
	}
	await Task.WhenAll(TaskList.ToArray());

}
async Task ProcessExtension(string File, String Destination, String Extension)
{

	string _tagetExtension = Extension.Replace(".", "");
	System.IO.FileInfo _fi = new System.IO.FileInfo(File);
	if (_fi.Extension == null || _fi.Extension.Length <= 0)// || _fi.Extension.Replace(".","") != _tagetExtension)
	{
		String _Destionation = System.IO.Path.Combine(Destination, $"{_fi.Name}.{_tagetExtension}");
		int _rename = 0;
		while (System.IO.File.Exists(_Destionation))
		{
			_Destionation = System.IO.Path.Combine(Destination, $"{_fi.Name}_{_rename}.{_tagetExtension}");
			_rename += 1;
		}
		System.IO.File.Move(_fi.FullName, _Destionation);
		Console.WriteLine($"Processing file {File} - {_Destionation}");
	}
}
#endregion


#region FileMove

void StartMove()
{
	string _target = getSourceFolder();
	string _destination = getDestinationFolder();

	Console.WriteLine($"Going from {_target} to {_destination}");

	processDirectory(_target, _destination);

	showMenu();

}
void processDirectory(string Path, String Destination)
{
	Console.WriteLine($"Processing directory {Path}");
	if (System.IO.Directory.Exists(Path))
	{
		foreach(string f in System.IO.Directory.GetFiles(Path))
		{
			ProcessFile(f, Destination);
		}
		if (System.IO.Directory.GetDirectories(Path).Length > 0)
		{
			foreach(var v in System.IO.Directory.GetDirectories(Path))
			{
				processDirectory(v, Destination);
			}
		}
	}
}
void ProcessFile(String FilePath, String Destination)
{
	Console.WriteLine($"Processing file {FilePath}");
	if (System.IO.File.Exists(FilePath))
	{
		String _fileName = System.IO.Path.GetFileName(FilePath);
		String _destinationPath = System.IO.Path.Combine(Destination, _fileName);
		int _rename = 0;
		while(System.IO.File.Exists(_destinationPath))
		{
			_destinationPath = System.IO.Path.Combine(Destination, $"{_rename}_{_fileName}");
			_rename += 1;
		}
		try
		{
			Console.WriteLine($"Move {_fileName} to {_destinationPath}");
			System.IO.File.Move(FilePath, _destinationPath);
		}
		catch(System.Exception e) {
			Console.WriteLine(e.Message);

		}
	}
}
#endregion

#region Email Process

async Task StartEmail(bool skipDuplicate)
{
	Console.WriteLine("Processing Email");
	string _target = getSourceFolder();
	string _destination = getDestinationFolder();

	Console.WriteLine($"Going from {_target} to {_destination}");
	await processDirectoryMails(_target, _destination, skipDuplicate);
}

async Task processDirectoryMails(string Path, String Destination, bool skipDuplicate)
{
	List<Task> TaskList = new List<Task>();
	Console.WriteLine($"Processing directory {Path}");
	if (System.IO.Directory.Exists(Path))
	{
		foreach (string f in System.IO.Directory.GetFiles(Path))
		{
			if(f.EndsWith(".eml"))
			{
				var LastTask = Task.Run(async() => await ProcessEmailFile(f, Destination, skipDuplicate));
				TaskList.Add(LastTask);
			}
		}
		if (System.IO.Directory.GetDirectories(Path).Length > 0)
		{
			foreach (var v in System.IO.Directory.GetDirectories(Path))
			{
				await processDirectoryMails(v, Destination, skipDuplicate);
			}
		}
	}
	await Task.WhenAll(TaskList.ToArray());
}
async Task ProcessEmailFile(String FilePath, String Destination, bool skipDuplicate)
{
	if (System.IO.File.Exists(FilePath))
	{
		var message = MimeMessage.Load(FilePath);
		if (message != null)
		{
			await ProcessEmail(message, Destination, skipDuplicate);
		}
	}
}
async Task ProcessEmail(MimeMessage message, String Destination, bool skipDuplicate)
{ 
	if(message != null)
	{
		foreach (var v in message.BodyParts)
		{
			try
			{
				string _fileName = ((MimeKit.MimePart)v).FileName;
				byte[] _fileContents = null;

				if (_fileName != null && _fileName.Length > 0)
				{

					using (var ms = new MemoryStream())
					{
						if (v is MessagePart)
						{
							var part = (MessagePart)v;
							await part.Message.WriteToAsync(ms);
						}
						if (v is MimePart)
						{
							var part = (MimePart)v;
							await part.Content.DecodeToAsync(ms);
							_fileContents = ms.ToArray();
						}

						String _destinationPath = System.IO.Path.Combine(Destination, $"{_fileName}");
						if (!skipDuplicate)
						{
							int _rename = 0;
							while (System.IO.File.Exists(_destinationPath))
							{
								_destinationPath = System.IO.Path.Combine(Destination, $"{_rename}_{_fileName}");
								_rename += 1;
							}
						}
						if (!System.IO.File.Exists(_destinationPath))
						{
							if (_fileContents != null)
							{
								Console.WriteLine($"Processing email {_fileName}");
								using (FileStream fs = new FileStream(_destinationPath, FileMode.OpenOrCreate))
								{

									await fs.WriteAsync(_fileContents, 0, _fileContents.Length);
									fs.Flush();
								}
							}
						}
						else
						{
							Console.WriteLine($"skipped email {_fileName}");
						}
					}
				}
			}
			catch (System.Exception e)
			{
				Console.WriteLine($"[ERROR] : {e.Message}");
			}
		}
	}
}
#endregion

#region Mbox
async Task StartMboxes()
{
	string _target = getSourceFolder();
	string _destination = getDestinationFolder();

	Console.WriteLine($"Going from {_target} to {_destination}");
	processDirectoryMboxes(_target, _destination, true);

}
async Task processDirectoryMboxes(string Path, String Destination, bool skipDuplicate)
{
	List<Task> TaskList = new List<Task>();
	Console.WriteLine($"Processing directory {Path}");
	if (System.IO.Directory.Exists(Path))
	{
		foreach (string f in System.IO.Directory.GetFiles(Path))
		{
			if (f.EndsWith(".mbox"))
			{
				var LastTask = Task.Run(async () => await ProcessMboxFile(f, Destination, skipDuplicate));
				TaskList.Add(LastTask);
			}
		}
		if (System.IO.Directory.GetDirectories(Path).Length > 0)
		{
			foreach (var v in System.IO.Directory.GetDirectories(Path))
			{
				await processDirectoryMboxes(v, Destination, skipDuplicate);
			}
		}
	}
	await Task.WhenAll(TaskList.ToArray());
}

async Task ProcessMboxFile(string FilePath, string DestionationPath, bool skipDuplicate)
{
	if (System.IO.File.Exists(FilePath))
	{
		using (System.IO.FileStream fs = new FileStream(FilePath, FileMode.Open))
		{

			var parser = new MimeParser(fs, MimeFormat.Mbox);
			while (!parser.IsEndOfStream)
			{
				MimeMessage _mboxMessage = parser.ParseMessage();
				if (_mboxMessage != null)
				{
					await ProcessEmail(_mboxMessage, DestionationPath, skipDuplicate);
				}
			}
		}
	}
}
#endregion


showMenu();