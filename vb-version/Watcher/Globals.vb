Friend Module Globals

    '--- are we running interactively or not
    Public Interactive As Boolean = False

    ''--- application settings
    'Public AppSettings As ApplicationSettings

    '--- the trigger info xml filename
    Public TriggerInfoFilename As String

    '--- the XML document defining the App settings
    Public AppConfigXml As Xml.XmlDocument

    '--- the XML document defining the triggers to be watched for
    'Public TriggerConfigXml As Xml.XmlDocument

    '--- the parser will contain the jobdefs
    Public Parser As JobDefinitionParser

    '--- collections for the created Watchers
    Public Watchers As Hashtable
    Public FileFolderWatchers As Hashtable

    '--- collection of the Triggers that will be iterated through
    'Public TriggerCollection As Hashtable

    '--- the collection of JobGroups
    Public JobGroupList As List(Of JobGroup)

    '--- the poller object that controls the polling of the triggers
    Public Manager As WatcherManager

End Module
