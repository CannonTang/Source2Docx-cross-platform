using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Source2Docx.Models;

internal sealed class SourceFileItem : INotifyPropertyChanged
{
	private bool isChecked;

	private string name = string.Empty;

	private string fullPath = string.Empty;

	public event PropertyChangedEventHandler PropertyChanged;

	public bool IsChecked
	{
		get { return isChecked; }
		set
		{
			if (isChecked == value)
			{
				return;
			}

			isChecked = value;
			OnPropertyChanged();
		}
	}

	public string Name
	{
		get { return name; }
		set
		{
			if (name == value)
			{
				return;
			}

			name = value;
			OnPropertyChanged();
		}
	}

	public string FullPath
	{
		get { return fullPath; }
		set
		{
			if (fullPath == value)
			{
				return;
			}

			fullPath = value;
			OnPropertyChanged();
		}
	}

	private void OnPropertyChanged([CallerMemberName] string propertyName = null)
	{
		PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
	}
}
