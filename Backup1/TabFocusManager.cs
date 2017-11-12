// Copyright (c) Microsoft Corporation.  All rights reserved.
using Accessibility;
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;

namespace Word2007RedactionTool
{
    /// <summary>
    /// Moves tab focus to the specified tab.
    /// </summary>
    class TabFocusManager
    {
        private const string AppName = "OpusApp";
        private const string RibbonName = "Ribbon";

        private string TargetControl;
        private IntPtr AppWindow;
        private IntPtr RibbonWindow;

        internal TabFocusManager()
        {
            AppWindow = GetTopLevelWindow(AppName, string.Empty);
            if (AppWindow != IntPtr.Zero)
                GetChildWindow(AppWindow);
            else
                System.Diagnostics.Debug.Fail("failed to get app window, aborting");
        }

        private static IntPtr GetTopLevelWindow(string AppName, string Caption)
        {
            if (string.IsNullOrEmpty(AppName) && string.IsNullOrEmpty(Caption))
                throw new ArgumentException("must supply one argument");
            else if (string.IsNullOrEmpty(AppName))
                return NativeMethods.FindWindowByCaption(IntPtr.Zero, Caption);
            else if (string.IsNullOrEmpty(Caption))
                return NativeMethods.FindWindowByClass(AppName, IntPtr.Zero);
            else
                return NativeMethods.FindWindow(AppName, Caption);
        }

        private void GetChildWindow(IntPtr Window)
        {
            NativeMethods.EnumWindowsDelegate EnumDelegate = new NativeMethods.EnumWindowsDelegate(ChildWindowFound);
            if (NativeMethods.EnumChildWindows(Window, EnumDelegate, IntPtr.Zero) < 0)
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Enumerating child windows failed.");
        }

        private int ChildWindowFound(IntPtr Window, IntPtr Params)
        {
            if (RibbonWindow == IntPtr.Zero)
            {
                // get the text from the window
                StringBuilder bld = new StringBuilder(256);

                if (NativeMethods.GetWindowText(Window, bld, 256) < 0)
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Getting window text failed.");

                string text = bld.ToString();

                System.Diagnostics.Debug.WriteLine("checking child window: " + text);

                if (text.ToString() == RibbonName)
                {
                    RibbonWindow = Window;
                    return 0;
                }                
            }
            return 1;
        }

        internal void Execute(string Target)
        {
            TargetControl = Target;

            IAccessible WordAcc = GetAccessibleObject(RibbonWindow);
            IAccessible RibbonAcc = GetObjectByName(WordAcc, TargetControl);
            if (RibbonAcc != null)
                RibbonAcc.accDoDefaultAction(0);
            else
                System.Diagnostics.Debug.Fail("ribbon switching failed");
        }

        private static IAccessible GetAccessibleObject(IntPtr Window)
        {
            Guid GuidIAccessible = new Guid("{618736E0-3C3D-11CF-810C-00AA00389B71}");
            return ((IAccessible)NativeMethods.AccessibleObjectFromWindow(Window, 0, ref GuidIAccessible));
        }

        private IAccessible GetObjectByName(object Control, string ChildControl)
        {
            object[] Children;
            IAccessible Result = null;

            IAccessible Parent = Control as IAccessible;
            if (Parent != null)
            {
                System.Diagnostics.Debug.WriteLine(Parent.get_accName(0));
                if (Parent.get_accName(0) == ChildControl)
                {
                    Result = Parent;
                }
                else
                {
                    Children = GetAccessibleChildren(Parent);
                    for (long i = 0; i <= Children.Length - 1; i++)
                    {
                        Result = GetObjectByName(Children[i], ChildControl);
                        if (Result != null)
                            break;
                    }
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("control skipped for being null");
            }
            return Result;
        }

        private static object[] GetAccessibleChildren(IAccessible Parent)
        {
            object[] Children = new Object[Parent.accChildCount];
            int obtained = 0;
            if (NativeMethods.AccessibleChildren(Parent, 0, Parent.accChildCount, Children, out obtained) < 0)
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Getting Accessible children failed.");
            System.Diagnostics.Debug.WriteLine("number of accessible children: " + Children.Length);
            return Children;
        }
    }

    /// <summary>
    /// Native methods needed by the TabFocusManager.
    /// </summary>
    class NativeMethods
    {
        private NativeMethods()
        {}

        #region OLEAcc Functions

        [DllImport("oleacc.dll", PreserveSig = false)]
        [return: MarshalAs(UnmanagedType.Interface)]
        public static extern object AccessibleObjectFromWindow(IntPtr Window, uint dwId, ref Guid GuidIAccessible);

        [DllImport("oleacc.dll", SetLastError = true)]
        public static extern uint AccessibleChildren(IAccessible paccContainer, int iChildStart, int cChildren, [Out] object[] rgvarChildren, out int pcObtained);

        #endregion

        #region Win32 Functions

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindow(string ClassName, string WindowName);

        [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindowByClass(string ClassName, IntPtr Nothing);

        [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindowByCaption(IntPtr Nothing, string Caption);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int EnumChildWindows(IntPtr Parent, EnumWindowsDelegate EnumFunction, IntPtr Parameters);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetWindowText(IntPtr Window, StringBuilder Text, int Length);

        [DllImport("user32.dll", CharSet =  CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, IntPtr windowTitle);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        #endregion

        public delegate int EnumWindowsDelegate(IntPtr Window, IntPtr Parameter);
    }
}
