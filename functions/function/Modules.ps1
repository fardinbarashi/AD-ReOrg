<#
    Loads everything the app depends on at runtime:
      * the .NET/WPF assemblies used to build the window from XAML
      * the RSAT ActiveDirectory module
      * the ReOrg.GroupPlan, ReOrg.RestoreRow and ReOrg.RenameRow bindable row
        types (the create-groups preview grid, the restore-lost-groups grid and
        the rename-groups grid)
    Kept in its own file - dot-sourced like the other function files - and run
    once from AD-ReOrg.ps1 via Import-AppModules before the window is shown.
#>

function Import-AppModules {
    # --- .NET/WPF assemblies ---------------------------------------------
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Xaml

    # --- RSAT ActiveDirectory module -------------------------------------
    Import-Module ActiveDirectory -ErrorAction continue

    # --- Bindable row types for the editable grids ------------------------
    # Each type is compiled behind its own type check. A .NET type can never be
    # unloaded or replaced inside a running PowerShell process, so a console
    # that already holds an older ReOrg assembly (one that knew GroupPlan but
    # not RestoreRow) must still be able to add the missing type.
    if (-not ('ReOrg.GroupPlan' -as [type])) {
        Add-Type -TypeDefinition @'
using System.ComponentModel;
namespace ReOrg {
    // One source group -> one new group on the "Groups : Create Groups" tab.
    public class GroupPlan : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnChanged(string n) {
            PropertyChangedEventHandler h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(n));
        }
        public bool   Include    { get; set; }
        public string SourceName { get; set; }
        public string SourceSam  { get; set; }
        public string SourceDN   { get; set; }
        public int    MemberCount{ get; set; }

        private string _newName;
        public string NewName { get { return _newName; } set { _newName = value; OnChanged("NewName"); } }

        private string _newSam;
        public string NewSam  { get { return _newSam; }  set { _newSam = value;  OnChanged("NewSam"); } }

        private string _status = "";
        public string Status  { get { return _status; }  set { _status = value;  OnChanged("Status"); } }
    }
}
'@
    }

    if (-not ('ReOrg.RestoreRow' -as [type])) {
        Add-Type -TypeDefinition @'
using System.ComponentModel;
namespace ReOrg {
    // One user + one lost group on the "Users : Restore lost groups" tab.
    public class RestoreRow : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnChanged(string n) {
            PropertyChangedEventHandler h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(n));
        }
        public string Account { get; set; }
        public string Name    { get; set; }
        public string Group   { get; set; }
        public string UserDN  { get; set; }
        public string GroupDN { get; set; }

        private bool _include = true;
        public bool Include { get { return _include; } set { _include = value; OnChanged("Include"); } }

        private string _rowStatus = "";
        public string Status { get { return _rowStatus; } set { _rowStatus = value; OnChanged("Status"); } }
    }
}
'@
    }

    if (-not ('ReOrg.RenameRow' -as [type])) {
        Add-Type -TypeDefinition @'
using System.ComponentModel;
namespace ReOrg {
    // One group on the "Groups : Rename Groups" tab: current and new values.
    public class RenameRow : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnChanged(string n) {
            PropertyChangedEventHandler h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(n));
        }
        public string ObjectGuid { get; set; }
        public string OldDN      { get; set; }
        public string NewDN      { get; set; }
        public string OldName    { get; set; }
        public string OldSam     { get; set; }
        public string OldDisplay { get; set; }
        public string OldDesc    { get; set; }

        private bool _include = true;
        public bool Include { get { return _include; } set { _include = value; OnChanged("Include"); } }

        private string _newName = "";
        public string NewName { get { return _newName; } set { _newName = value; OnChanged("NewName"); } }
        private string _newSam = "";
        public string NewSam { get { return _newSam; } set { _newSam = value; OnChanged("NewSam"); } }
        private string _newDisplay = "";
        public string NewDisplay { get { return _newDisplay; } set { _newDisplay = value; OnChanged("NewDisplay"); } }
        private string _newDesc = "";
        public string NewDesc { get { return _newDesc; } set { _newDesc = value; OnChanged("NewDesc"); } }

        private string _rowStatus = "";
        public string Status { get { return _rowStatus; } set { _rowStatus = value; OnChanged("Status"); } }
    }
}
'@
    }
}
