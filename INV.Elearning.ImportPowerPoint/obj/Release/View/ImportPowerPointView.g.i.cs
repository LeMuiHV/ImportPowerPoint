﻿#pragma checksum "..\..\..\View\ImportPowerPointView.xaml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "FDCBE1145FB3098B383707B875D973D9FC8FCE94"
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using INV.Elearning.Controls;
using INV.Elearning.ImportPowerPoint.Converter;
using INV.Elearning.ImportPowerPoint.View;
using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Media.TextFormatting;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Shell;


namespace INV.Elearning.ImportPowerPoint.View {
    
    
    /// <summary>
    /// ImportPowerPointView
    /// </summary>
    public partial class ImportPowerPointView : INV.Elearning.Controls.ElearningWindow, System.Windows.Markup.IComponentConnector {
        
        
        #line 27 "..\..\..\View\ImportPowerPointView.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.RadioButton rdBackground;
        
        #line default
        #line hidden
        
        
        #line 28 "..\..\..\View\ImportPowerPointView.xaml"
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        internal System.Windows.Controls.RadioButton rdObjInSlide;
        
        #line default
        #line hidden
        
        private bool _contentLoaded;
        
        /// <summary>
        /// InitializeComponent
        /// </summary>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        public void InitializeComponent() {
            if (_contentLoaded) {
                return;
            }
            _contentLoaded = true;
            System.Uri resourceLocater = new System.Uri("/INV.Elearning.ImportPowerPoint;component/view/importpowerpointview.xaml", System.UriKind.Relative);
            
            #line 1 "..\..\..\View\ImportPowerPointView.xaml"
            System.Windows.Application.LoadComponent(this, resourceLocater);
            
            #line default
            #line hidden
        }
        
        [System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [System.CodeDom.Compiler.GeneratedCodeAttribute("PresentationBuildTasks", "4.0.0.0")]
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Never)]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity")]
        [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily")]
        void System.Windows.Markup.IComponentConnector.Connect(int connectionId, object target) {
            switch (connectionId)
            {
            case 1:
            
            #line 11 "..\..\..\View\ImportPowerPointView.xaml"
            ((INV.Elearning.ImportPowerPoint.View.ImportPowerPointView)(target)).Closed += new System.EventHandler(this.ElearningWindow_Closed);
            
            #line default
            #line hidden
            return;
            case 2:
            this.rdBackground = ((System.Windows.Controls.RadioButton)(target));
            return;
            case 3:
            this.rdObjInSlide = ((System.Windows.Controls.RadioButton)(target));
            return;
            case 4:
            
            #line 90 "..\..\..\View\ImportPowerPointView.xaml"
            ((System.Windows.Controls.Button)(target)).Click += new System.Windows.RoutedEventHandler(this.OK_Click);
            
            #line default
            #line hidden
            return;
            }
            this._contentLoaded = true;
        }
    }
}

