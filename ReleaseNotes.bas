Attribute VB_Name = "ReleaseNotes"
Option Explicit

Public Function ReleaseNotesInfo() As String

  Dim R As String
  
  R = "{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 MS Sans Serif;}{\f1\fnil\fcharset2 Symbol;}}"
  R = R & "{\colortbl ;\red128\green0\blue128;\red128\green0\blue0;\red0\green0\blue128;\red255\green0\blue0;}"
  R = R & "\viewkind4\uc1\pard\cf1\lang1033\b\f0\fs20 Source Search 2.0"
  R = R & "\par Copyright 2001\cf2 "
  R = R & "\par \cf3 Created by Casey Goodhew"
  R = R & "\par goodhewc@hotmail.com"
  R = R & "\par http://goodhew.2y.net"
  R = R & "\par \cf0\b0 "
  R = R & "\par \cf1\b Source Search \cf0\b0 now uses code from:"
  R = R & "\par \b Richsoft Computing www.richsoftcomputing.btinternet.co.uk\b0 "
  R = R & "\par Thanks for help!"
  R = R & "\par "
  R = R & "\par \cf4\b Don't forget to copy:"
  R = R & "\par "
  R = R & "\par \tab Unzdll.dll"
  R = R & "\par \tab Zipdll.dll"
  R = R & "\par \tab Zipit.dll"
  R = R & "\par "
  R = R & "\par to your ...\\Windows\\System32\\ directory."
  R = R & "\par \cf0\b0 "
  R = R & "\par "
  R = R & "\par \cf3\ul\b VERSION 2.0"
  R = R & "\par \cf0\ulnone\b0 "
  R = R & "\par \pard{\pntext\f1\'B7\tab}{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\'B7}}\fi-720\li720 Added functionality for dragging and dropping ZIP files with VB project contents to the application."
  R = R & "\par {\pntext\f1\'B7\tab}Added node info box that gives a summary of a node in the project tree."
  R = R & "\par {\pntext\f1\'B7\tab}Added Popup menu to project tree. Just right click away!"
  R = R & "\par {\pntext\f1\'B7\tab}Updated the First Run Wizard so that it will handle migration to the new dictionary and storeage path in the registry."
  R = R & "\par {\pntext\f1\'B7\tab}\pard "
  R = R & "\par "
  R = R & "\par \cf3\ul\b VERSION 1.1"
  R = R & "\par \cf0\ulnone\b0 "
  R = R & "\par \pard{\pntext\f1\'B7\tab}{\*\pn\pnlvlblt\pnf1\pnindent0{\pntxtb\'B7}}\fi-720\li720 Added 'Ignore Comments Option'"
  R = R & "\par {\pntext\f1\'B7\tab}\pard "
  R = R & "\par "
  R = R & "\par \cf3\ul\b VERSION 1.0"
  R = R & "\par \cf0\ulnone\b0 "
  R = R & "\par If this is Version 1.0 ..... isn't everything new?\fs16 "
  R = R & "\par }"
  R = R & ""


  
  ReleaseNotesInfo = R

End Function


