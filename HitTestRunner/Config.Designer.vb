'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.239
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.Desktop.AddIns
Imports ESRI.ArcGIS.Framework
Imports System
Imports System.Collections.Generic

Namespace My
    
    '''<summary>
    '''A class for looking up declarative information in the associated configuration xml file (.esriaddinx).
    '''</summary>
    Friend Module ThisAddIn
        
        Friend ReadOnly Property Name() As String
            Get
                Return "HitTestRunner"
            End Get
        End Property
        
        Friend ReadOnly Property AddInID() As String
            Get
                Return "{435f245e-5b13-4bbf-8efc-cd5d4f84c5c6}"
            End Get
        End Property
        
        Friend ReadOnly Property Company() As String
            Get
                Return "WDFW"
            End Get
        End Property
        
        Friend ReadOnly Property Version() As String
            Get
                Return "1.0"
            End Get
        End Property
        
        Friend ReadOnly Property Description() As String
            Get
                Return "Type in a description for this Add-in."
            End Get
        End Property
        
        Friend ReadOnly Property Author() As String
            Get
                Return "pierckbp"
            End Get
        End Property
        
        Friend ReadOnly Property [Date]() As String
            Get
                Return "05/03/2012"
            End Get
        End Property
        
        '''<summary>
        '''A class for looking up Add-in id strings declared in the associated configuration xml file (.esriaddinx).
        '''</summary>
        Friend Class IDs
            
            '''<summary>
            '''Returns 'WDFW_HitTestRunner_HitTestRunner', the id declared for Add-in Button class 'HitTestRunner'
            '''</summary>
            Friend Shared ReadOnly Property HitTestRunner() As String
                Get
                    Return "WDFW_HitTestRunner_HitTestRunner"
                End Get
            End Property
        End Class
    End Module
    
Friend Module ArcMap
  Private s_app As ESRI.ArcGIS.Framework.IApplication
  Private s_docEvent As ESRI.ArcGIS.ArcMapUI.IDocumentEvents_Event

  Public ReadOnly Property Application() As ESRI.ArcGIS.Framework.IApplication
    Get
      If s_app Is Nothing Then
        s_app = TryCast(Internal.AddInStartupObject.GetHook(Of ESRI.ArcGIS.ArcMapUI.IMxApplication)(), ESRI.ArcGIS.Framework.IApplication)
      End If

      Return s_app
    End Get
  End Property

  Public ReadOnly Property Document() As ESRI.ArcGIS.ArcMapUI.IMxDocument
    Get
      If Application IsNot Nothing Then
        Return TryCast(Application.Document, ESRI.ArcGIS.ArcMapUI.IMxDocument)
      End If

      Return Nothing
    End Get
  End Property
  Public ReadOnly Property ThisApplication() As ESRI.ArcGIS.ArcMapUI.IMxApplication
    Get
      Return TryCast(Application, ESRI.ArcGIS.ArcMapUI.IMxApplication)
    End Get
  End Property
  Public ReadOnly Property DockableWindowManager() As ESRI.ArcGIS.Framework.IDockableWindowManager
    Get
      Return TryCast(Application, ESRI.ArcGIS.Framework.IDockableWindowManager)
    End Get
  End Property

  Public ReadOnly Property Events() As ESRI.ArcGIS.ArcMapUI.IDocumentEvents_Event
    Get
      s_docEvent = TryCast(Document, ESRI.ArcGIS.ArcMapUI.IDocumentEvents_Event)
      Return s_docEvent
    End Get
  End Property
End Module

Namespace Internal
  <ESRI.ArcGIS.Desktop.AddIns.StartupObjectAttribute(), _
   Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), _
   Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()> _
  Partial Public Class AddInStartupObject
    Inherits ESRI.ArcGIS.Desktop.AddIns.AddInEntryPoint

    Private m_addinHooks As List(Of Object)
    Private Shared _sAddInHostManager As AddInStartupObject

    Public Sub New()

    End Sub

    <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Protected Overrides Function Initialize(ByVal hook As Object) As Boolean
      Dim createSingleton As Boolean = _sAddInHostManager Is Nothing
      If createSingleton Then
        _sAddInHostManager = Me
        m_addinHooks = New List(Of Object)
        m_addinHooks.Add(hook)
      ElseIf Not _sAddInHostManager.m_addinHooks.Contains(hook) Then
        _sAddInHostManager.m_addinHooks.Add(hook)
      End If

      Return createSingleton
    End Function

    <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Protected Overrides Sub Shutdown()
      _sAddInHostManager = Nothing
      m_addinHooks = Nothing
    End Sub

    <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Friend Shared Function GetHook(Of T As Class)() As T
      If _sAddInHostManager IsNot Nothing Then
        For Each o As Object In _sAddInHostManager.m_addinHooks
          If TypeOf o Is T Then
            Return DirectCast(o, T)
          End If
        Next
      End If

      Return Nothing
    End Function

    ''' <summary>
    ''' Expose this instance of Add-in class externally
    ''' </summary>
    Public Shared Function GetThis() As AddInStartupObject
      Return _sAddInHostManager
    End Function

  End Class
End Namespace

End Namespace
