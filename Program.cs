using MimeKit;
using System.Data.SQLite;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

#region Prompts & text

string _promptSourceFolder = "Enter source folder:";
string _promptSourceFolderNotFound = "Source folder not found, enter valid source folder:";
string _promptFolderNotFound = "Folder not found, enter valid path:";
string _promptDestinationFolder = "Enter destination folder:";
string _promptDestionationFolderNotFound = "Destination folder not found, enter valid source folder:";
string _titleStartExtentions = "Processing files to set extensions";
string _promptExtension = "Enter extension for files:";
string _promptExtensionIsEmpty = "Please give an extension";
string _promptDatabaseFolder = "Enter path for database file:";
string _promptDatabaseName = "Enter name for database file:";


string getDatabaseFolder()
{
    Console.WriteLine();
    Console.WriteLine(_promptDatabaseFolder);
    string? _target = Console.ReadLine();
    while (!System.IO.Directory.Exists(_target))
    {
        Console.WriteLine(_promptFolderNotFound);
        _target = getDatabaseFolder();
    }
    return _target;
}
string getDatabaseName()
{
    Console.WriteLine();
    Console.WriteLine(_promptDatabaseName);
    string? _target = Console.ReadLine();
    while (_target == null || _target.Length<=0)
    {
        Console.WriteLine(_promptDatabaseName);
        _target = getDatabaseName();
    }
    return _target;
}
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
    Console.WriteLine("Set [D]atabase file");
    Console.WriteLine("[E]mail");
	Console.WriteLine("m[B]oxes");
	Console.WriteLine("[M]ove to flat folder structure");
	Console.WriteLine("Exten[S]ion[S]");
	Console.WriteLine("e[X]it");
	ConsoleKeyInfo _input = Console.ReadKey();
	switch (_input.Key.ToString().ToUpper())
	{
		case "D":
			StartCreateDatabase();
			showMenu();
			break;
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


#region Database process

async Task StartCreateDatabase()
{
    Console.WriteLine("Creating Database");
    string _target = getDatabaseFolder();
    string _databaseName = getDatabaseName();

    Console.WriteLine($"Creating {_databaseName} at path {_target}");
    await processStartCreateDatabase(_target, _databaseName);
}

string _databasePath = "";

async Task processStartCreateDatabase(string Path, String filename)
{
	string _dbPath = System.IO.Path.Combine(Path, filename);
	// this creates a zero-byte file
	SQLiteConnection.CreateFile(_dbPath);	
	try
	{
		string connectionString = $"Data Source={_dbPath};Version=3;";
		using (SQLiteConnection m_dbConnection = new SQLiteConnection(connectionString))
		{
			m_dbConnection.Open();
			string sql = @"CREATE TABLE IF NOT EXISTS ""Email"" (
									""id""	INTEGER NOT NULL UNIQUE,
									""MigrationDate""	datetime,
									""MailDate""	datetime,
									""Subject""	TEXT,
									""HtmlBody""	TEXT,
									""TextBody""	TEXT,
									""From""	TEXT,
									""To""	TEXT,
									""CC""	TEXT,
									""BCC""	TEXT,
									PRIMARY KEY(""id"" AUTOINCREMENT));";
			SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
			command.ExecuteNonQuery();
			sql = @"CREATE TABLE IF NOT EXISTS ""Attachments"" (
									""id""	INTEGER NOT NULL UNIQUE,
									""MailId""	INTEGER NOT NULL,
									""MigrationDate""	datetime,
									""DestinationPath""	TEXT,
									""FileName""	TEXT,
									""OriginalFileName""	TEXT,
									PRIMARY KEY(""Id""),
									CONSTRAINT ""FK_EMAILID"" FOREIGN KEY(""MailId"") REFERENCES ""Email""(""id"")
								);";
            command = new SQLiteCommand(sql, m_dbConnection);
            command.ExecuteNonQuery();

            m_dbConnection.Close();
		}
        _databasePath = _dbPath;
    }
	catch(System.Exception ex)
	{
		Console.WriteLine($"[ERROR] {ex.Message}");
	}
	await Task.CompletedTask;
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
		long _mailID = 0;

        if (System.IO.Path.Exists(_databasePath))
		{
			_mailID = await SaveMessageToDB(_databasePath,message);
		}
		foreach (var v in message.BodyParts)
		{
			try
			{
				if( v is MessagePart)
				{
					MessagePart part = (MessagePart)v;
					if(part != null) // this is an attachment that is also an email
					{
						await ProcessEmail(part.Message, Destination, skipDuplicate);
					}
				}
				else if (v is MimePart && ((MimeKit.MimePart)v).IsAttachment)
				{

					string _fileName = ((MimeKit.MimePart)v).FileName;
                    if (_fileName!=null && _fileName.Length>0)
                    {
						_fileName = SanitizeFileName(_fileName);
					}
					else
					{

					}

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
							if(_mailID>0)
							{
								await SaveAttachmentToDB(_databasePath, _mailID, Destination, _fileName, System.IO.Path.GetFileName(_destinationPath));
							}
						}
					}
				}
				else if (v.IsAttachment)
				{
				}
				else
				{
					// just text 
				}
			}
			catch (System.Exception e)
			{
				Console.WriteLine($"[ERROR] : {e.Message}");
			}
		}
	}
}
async Task<long> SaveAttachmentToDB(String Path, long mailID, string destinationPath,string OriginalFileName, string FileName)
{
    long RowID = 0;
    try
    {
        string connectionString = $"Data Source={Path};Version=3;";
        using (SQLiteConnection m_dbConnection = new SQLiteConnection(connectionString))
        {
            m_dbConnection.Open();
            string sql = $@"INSERT INTO Attachments ('MailId', 'MigrationDate', 'DestinationPath', 'OriginalFileName', 'FileName')
						VALUES (@mailID, @migrationdate, @destinationpath, @originalfilename, @filename) RETURNING id";
            SQLiteCommand cmd = new SQLiteCommand(sql, m_dbConnection);
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@migrationdate", DateTime.Now);
            cmd.Parameters.AddWithValue("@mailID", mailID);
            cmd.Parameters.AddWithValue("@destinationpath", destinationPath);
            cmd.Parameters.AddWithValue("@originalfilename", OriginalFileName);
            cmd.Parameters.AddWithValue("@filename", FileName);

            RowID = (long)cmd.ExecuteScalar();
            m_dbConnection.Close();
        }
    }
    catch (System.Exception ex)
    {
        Console.WriteLine($"[ERROR] {ex.Message}");
    }
    return RowID;
}
async Task<long> SaveMessageToDB(String Path, MimeMessage message)
{
	long RowID = 0;
    try
    {
        string connectionString = $"Data Source={Path};Version=3;";
        using (SQLiteConnection m_dbConnection = new SQLiteConnection(connectionString))
        {
			m_dbConnection.Open();
			string sql = $@"INSERT INTO Email ('MigrationDate', 'MailDate', 'Subject', 'HtmlBody', 'TextBody', 'from', 'To', 'CC', 'BCC')
						VALUES (@migrationdate, @maildate, @subject, @htmlbody, @textbody, @mailfrom, @mailto,@mailcc, @mailbcc) RETURNING id;";
            SQLiteCommand cmd = new SQLiteCommand(sql, m_dbConnection);
            cmd.CommandText = sql;
            cmd.Parameters.AddWithValue("@migrationdate", DateTime.Now);
            cmd.Parameters.AddWithValue("@maildate", message.Date.DateTime);
            cmd.Parameters.AddWithValue("@subject", message.Subject);
            cmd.Parameters.AddWithValue("@htmlbody", message.HtmlBody);
            cmd.Parameters.AddWithValue("@textbody", message.TextBody);
            cmd.Parameters.AddWithValue("@mailfrom", string.Join(",", message.From));
            cmd.Parameters.AddWithValue("@mailto", string.Join(",", message.To));
            cmd.Parameters.AddWithValue("@mailcc", string.Join(",", message.Cc));
            cmd.Parameters.AddWithValue("@mailbcc", string.Join(",", message.Bcc));

            RowID = (long)cmd.ExecuteScalar();
            m_dbConnection.Close();
        }
    }
    catch (System.Exception ex)
    {
        Console.WriteLine($"[ERROR] {ex.Message}");
    }
	return RowID;
}
string SanitizeFileName(string fileName)
{
	string _cleanString = "";
	try
	{
		if(fileName != null && fileName.Length > 0)
		{
			// there can be attachmments with slashes or network paths as name. Which will throw off the destination path.
			_cleanString = Regex.Replace(fileName, "[^a-zA-Z0-9_.]+", "_", RegexOptions.Compiled);
		}
		if (_cleanString.Length <= 0)
		{
			string _extension = fileName != null ? System.IO.Path.GetExtension(fileName) : "";
			_cleanString = $"{DateTime.Now.ToString("yyyymmddHHmmssffff")}.{_extension}";
		}
	}
	catch(System.Exception e)
	{
		// something went really wrong
	}
	return _cleanString;
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