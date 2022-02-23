Imports System.IO
Imports OTIS

Public Class HashCreator

    Private _currentWatcher As WatcherBase
    Private _conjunctor As String


    '--- CONSTRUCTORS
    Sub New(ByVal CurrentWatcher As WatcherBase)

        _currentWatcher = CurrentWatcher

    End Sub


    '**********************************************************
    '* create the hash and return it
    '* 
    '* 
    '*
    '* SDCP 11-Oct-2006
    '**********************************************************
    Public Function GetHash() As String
        Dim builtHash As String = ""
        Dim hashItem As Xml.XmlNode


        '--- extract the hashbuilder node
        Dim hashBuilderNode As Xml.XmlNode = _currentWatcher.ActiveJob.Trigger.ExpandedTriggerInfoXml.SelectSingleNode("//hashbuilder")

        '--- get the conjunctor
        _conjunctor = CType(hashBuilderNode.SelectSingleNode("@conjunctor").Value, String)

        '--- get a SortedList of the hashitems
        Dim sortedHashItems As SortedList = SortNodeList(hashBuilderNode.SelectNodes("//hashitem"), "order", TypeCode.Int32)

        '--- loop through the list and attempt to build each item
        For Each hashItem In sortedHashItems.Values

            builtHash += hashItem.InnerText
            builtHash += _conjunctor

        Next

        '--- return the built hash
        Return builtHash

    End Function


End Class
