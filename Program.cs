using MimeKit;
using System.IO;
using System.Net.Mail;
using System.Text;

async void showMenu()
{
	Console.WriteLine("[E]mail");
	Console.WriteLine("[M]ove");
	Console.WriteLine("Exten[S]ion[S]");
	Console.WriteLine("e[X]it");
	ConsoleKeyInfo _input = Console.ReadKey();
	switch (_input.Key.ToString())
	{
		case "E": // E:\Email\Test
		case "e":
			StartEmail(true);
			showMenu();
			break;
		case "M":
		case "m":
			StartMove();
			break;
		case "S":
		case "s":
			StartExtensions();
			showMenu();
			break;
		case "X":
		case "x":
			Environment.Exit(0);
			break;
		default:
			showMenu();
			break;
	}
}
showMenu();


async Task StartExtensions()
{
	// E:\Email\Extension
	Console.WriteLine("Processing files to set extensions");

	Console.WriteLine("Enter extension for files:");
	var _extension = Console.ReadLine();

	Console.WriteLine("Enter source folder:");
	var _target = Console.ReadLine();
	while (!System.IO.Directory.Exists(_target))
	{

		Console.WriteLine("Source folder not found, enter valid source folder:");
		_target = Console.ReadLine();
	}
	Console.WriteLine("Enter destination folder:");

	var _destination = Console.ReadLine();
	while (!System.IO.Directory.Exists(_destination))
	{

		Console.WriteLine("Destination folder not found, enter valid source folder:");
		_destination = Console.ReadLine();
	}
	Console.WriteLine($"Going from {_target} to {_destination}");
	await processDirectoryExtensions(_target, _destination, _extension);
}

async Task StartEmail(bool skipDuplicate)
{
	Console.WriteLine("Processing Email");



	Console.WriteLine("Enter source folder:");
	var _target = Console.ReadLine();
	while (!System.IO.Directory.Exists(_target))
	{

		Console.WriteLine("Source folder not found, enter valid source folder:");
		_target = Console.ReadLine();
	}
	Console.WriteLine("Enter destination folder:");

	var _destination = Console.ReadLine();
	while (!System.IO.Directory.Exists(_destination))
	{

		Console.WriteLine("Destination folder not found, enter valid source folder:");
		_destination = Console.ReadLine();
	}
	Console.WriteLine($"Going from {_target} to {_destination}");
	await processDirectoryMails(_target, _destination, skipDuplicate);
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
	if (_fi.Extension == null || _fi.Extension.Length<=0)// || _fi.Extension.Replace(".","") != _tagetExtension)
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
void StartMove()
{

	Console.WriteLine("Enter source folder:");
	var _target = Console.ReadLine();
	while (!System.IO.Directory.Exists(_target))
	{

		Console.WriteLine("Source folder not found, enter valid source folder:");
		_target = Console.ReadLine();
	}
	Console.WriteLine("Enter destination folder:");
	var _destination = Console.ReadLine();
	while (!System.IO.Directory.Exists(_destination))
	{

		Console.WriteLine("Destination folder not found, enter valid source folder:");
		_destination = Console.ReadLine();
	}
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
				var LastTask = Task.Run(async() => await ProcessEmail(f, Destination, skipDuplicate));
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
async Task ProcessEmail(String FilePath, String Destination, bool skipDuplicate)
{
	if (System.IO.File.Exists(FilePath))
	{
		var message = MimeMessage.Load(FilePath);

		if (message != null)
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
									Console.WriteLine($"Processing email {FilePath}");
									using (FileStream fs = new FileStream(_destinationPath, FileMode.OpenOrCreate))
									{

										await fs.WriteAsync(_fileContents, 0, _fileContents.Length);
										fs.Flush();
									}
								}
							}
							else
							{
								Console.WriteLine($"skipped email {FilePath}");
							}
						}
					}
				}
				catch (System.Exception e)
				{
					Console.WriteLine($"[ERROR] {FilePath} : {e.Message}");
				}
			}
		}
	}
		
}