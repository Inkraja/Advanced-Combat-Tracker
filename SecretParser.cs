// reference:System.Core.dll
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Advanced_Combat_Tracker;
using System.Reflection;
using System.Xml;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Net;
using System.Diagnostics;
//Debug
// using System.Runtime.InteropServices;

// There is no copyright on this code

// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
// associated documentation files (the "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is furnished to do so.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
// LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN
// NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
// SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

[assembly: AssemblyTitle("Secret World damage and heal parse")]
[assembly: AssemblyDescription("Read through the CombatLog.txt files and parse the combat and healing done (ACT3)")]
[assembly: AssemblyCopyright("Author: Boorish, since 1.0.5.4 Lausi; Contributions from: Eafalas, Holok, Inkraja, Akamiko; ***")]
[assembly: AssemblyVersion("1.0.7.3")]
// This plugin is based on the Rift3 plugin by Creub and Altuslumen.  Thanks guys :)
// Fix for glance and penetrate hits fom Holok
// Added Incoming Damage (takencrit%, takenpen&, ...) to chat export (Holok)
// Add shiny colours to agis, bump version to 1.0.6.6 (Inkraja https://github.com/Inkraja/Advanced-Combat-Tracker)


namespace SecretParse_Plugin
{
    public class SecretParse : UserControl, IActPluginV1
    {
        // Debug
        //  [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        //  public static extern void OutputDebugString(string message);
        // /Debug

        #region Designer Created Code (Avoid editing)

        private System.ComponentModel.IContainer components = null;
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.labelHeader = new System.Windows.Forms.Label();
            this.labelGeneral = new System.Windows.Forms.Label();
            this.labelLanguage = new System.Windows.Forms.Label();
            this.comboBox_Language = new System.Windows.Forms.ComboBox();
            this.checkBox_AutoCheck = new System.Windows.Forms.CheckBox();
            this.checkBox_ExportScript = new System.Windows.Forms.CheckBox();
            this.checkBox_LimitNames = new System.Windows.Forms.CheckBox();
            this.checkBox_ExportRoundDPS = new System.Windows.Forms.CheckBox();
            this.checkBox_ExportShowLegend = new System.Windows.Forms.CheckBox();
            this.checkBox_ExportColored = new System.Windows.Forms.CheckBox();
            this.checkBox_ExportSplit = new System.Windows.Forms.CheckBox();
            this.checkBox_ExportHtml = new System.Windows.Forms.CheckBox();
            this.checkBox_ExportAllies = new System.Windows.Forms.CheckBox();
            this.checkBox_DontExportShortEnc = new System.Windows.Forms.CheckBox();
            this.checkBox_EnableTSWAddon = new System.Windows.Forms.CheckBox();
            this.checkBox_AutofixDBConf = new System.Windows.Forms.CheckBox();
            this.checkBox_HideUnknown = new System.Windows.Forms.CheckBox();
            this.checkBox_SelfDamage = new System.Windows.Forms.CheckBox();
            this.checkBox_SelfPlayerDamage = new System.Windows.Forms.CheckBox();
            this.checkBox_ReduceAegis = new System.Windows.Forms.CheckBox();
            this.labelFilterinig = new System.Windows.Forms.Label();
            this.checkBox_Filter = new System.Windows.Forms.CheckBox();
            this.labelFilteringName = new System.Windows.Forms.Label();
            this.textBox_FilterName = new System.Windows.Forms.TextBox();
            this.button_AddCharacter = new System.Windows.Forms.Button();
            this.button_DeleteCharacter = new System.Windows.Forms.Button();
            this.checkBox_filterExclude = new System.Windows.Forms.CheckBox();
            this.checkBox_filterScript = new System.Windows.Forms.CheckBox();
            this.labelChatExport = new System.Windows.Forms.Label();
            this.checkedListBox_ExportFields = new System.Windows.Forms.CheckedListBox();
            this.labelFilterChars = new System.Windows.Forms.Label();
            this.checkedListBox_Filters = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            //
            // labelHeader
            //
            this.labelHeader.AutoSize = true;
            this.labelHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelHeader.Location = new System.Drawing.Point(12, 9);
            this.labelHeader.Name = "labelHeader";
            this.labelHeader.Size = new System.Drawing.Size(214, 13);
            this.labelHeader.TabIndex = 26;
            this.labelHeader.Text = "Secret Combat parser plugin Options";
            this.labelHeader.MouseHover += new System.EventHandler(this.label1_MouseHover);
            //
            // labelGeneral
            //
            this.labelGeneral.AutoSize = true;
            this.labelGeneral.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelGeneral.Location = new System.Drawing.Point(12, 28);
            this.labelGeneral.Name = "labelGeneral";
            this.labelGeneral.Size = new System.Drawing.Size(44, 13);
            this.labelGeneral.TabIndex = 27;
            this.labelGeneral.Text = "General";
            //
            // labelLanguage
            //
            this.labelLanguage.AutoSize = true;
            this.labelLanguage.Location = new System.Drawing.Point(12, 44);
            this.labelLanguage.Name = "labelLanguage";
            this.labelLanguage.Size = new System.Drawing.Size(55, 13);
            this.labelLanguage.TabIndex = 28;
            this.labelLanguage.Text = "Language";
            //
            // comboBox_Language
            //
            this.comboBox_Language.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Language.FormattingEnabled = true;
            this.comboBox_Language.Items.AddRange(new object[] {
            "English",
            "German",
            "French"});
            this.comboBox_Language.Location = new System.Drawing.Point(73, 41);
            this.comboBox_Language.Name = "comboBox_Language";
            this.comboBox_Language.Size = new System.Drawing.Size(159, 21);
            this.comboBox_Language.TabIndex = 1;
            this.comboBox_Language.SelectedIndexChanged += new System.EventHandler(this.comboBox_Language_SelectedIndexChanged);
            this.comboBox_Language.MouseHover += new System.EventHandler(this.comboBox_Language_MouseHover);
            //
            // checkBox_AutoCheck
            //
            this.checkBox_AutoCheck.AutoSize = true;
            this.checkBox_AutoCheck.Checked = true;
            this.checkBox_AutoCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_AutoCheck.Location = new System.Drawing.Point(15, 62);
            this.checkBox_AutoCheck.Name = "checkBox_AutoCheck";
            this.checkBox_AutoCheck.Size = new System.Drawing.Size(320, 17);
            this.checkBox_AutoCheck.TabIndex = 2;
            this.checkBox_AutoCheck.Text = "Auto check the web page for new version of the Secret Plugin";
            this.checkBox_AutoCheck.UseVisualStyleBackColor = true;
            this.checkBox_AutoCheck.MouseHover += new System.EventHandler(this.checkBox_AutoCheck_MouseHover);
            //
            // checkBox_ExportScript
            //
            this.checkBox_ExportScript.AutoSize = true;
            this.checkBox_ExportScript.Checked = true;
            this.checkBox_ExportScript.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ExportScript.Location = new System.Drawing.Point(15, 78);
            this.checkBox_ExportScript.Name = "checkBox_ExportScript";
            this.checkBox_ExportScript.Size = new System.Drawing.Size(157, 17);
            this.checkBox_ExportScript.TabIndex = 3;
            this.checkBox_ExportScript.Text = "Export results to TSW script";
            this.checkBox_ExportScript.UseVisualStyleBackColor = true;
            this.checkBox_ExportScript.CheckedChanged += new System.EventHandler(this.checkBox_ExportScript_CheckedChanged);
            this.checkBox_ExportScript.MouseHover += new System.EventHandler(this.checkBox_ExportScript_MouseHover);
            //
            // checkBox_LimitNames
            //
            this.checkBox_LimitNames.AutoSize = true;
            this.checkBox_LimitNames.Checked = true;
            this.checkBox_LimitNames.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_LimitNames.Location = new System.Drawing.Point(35, 94);
            this.checkBox_LimitNames.Name = "checkBox_LimitNames";
            this.checkBox_LimitNames.Size = new System.Drawing.Size(159, 17);
            this.checkBox_LimitNames.TabIndex = 4;
            this.checkBox_LimitNames.Text = "Limit playernames to 7 chars";
            this.checkBox_LimitNames.UseVisualStyleBackColor = true;
            this.checkBox_LimitNames.MouseHover += new System.EventHandler(this.checkBox_LimitNames_MouseHover);
            //
            // checkBox_ExportRoundDPS
            //
            this.checkBox_ExportRoundDPS.AutoSize = true;
            this.checkBox_ExportRoundDPS.Location = new System.Drawing.Point(35, 110);
            this.checkBox_ExportRoundDPS.Name = "checkBox_ExportRoundDPS";
            this.checkBox_ExportRoundDPS.Size = new System.Drawing.Size(150, 17);
            this.checkBox_ExportRoundDPS.TabIndex = 5;
            this.checkBox_ExportRoundDPS.Text = "Round DPS / HPS values";
            this.checkBox_ExportRoundDPS.UseVisualStyleBackColor = true;
            this.checkBox_ExportRoundDPS.MouseHover += new System.EventHandler(this.checkBox_ExportRoundDPS_MouseHover);
            //
            // checkBox_ExportShowLegend
            //
            this.checkBox_ExportShowLegend.AutoSize = true;
            this.checkBox_ExportShowLegend.Checked = true;
            this.checkBox_ExportShowLegend.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ExportShowLegend.Location = new System.Drawing.Point(35, 126);
            this.checkBox_ExportShowLegend.Name = "checkBox_ExportShowLegend";
            this.checkBox_ExportShowLegend.Size = new System.Drawing.Size(155, 17);
            this.checkBox_ExportShowLegend.TabIndex = 6;
            this.checkBox_ExportShowLegend.Text = "Show legend in TSW script";
            this.checkBox_ExportShowLegend.UseVisualStyleBackColor = true;
            this.checkBox_ExportShowLegend.MouseHover += new System.EventHandler(this.checkBox_ExportShowLegend_MouseHover);
            //
            // checkBox_ExportColored
            //
            this.checkBox_ExportColored.AutoSize = true;
            this.checkBox_ExportColored.Location = new System.Drawing.Point(35, 142);
            this.checkBox_ExportColored.Name = "checkBox_ExportColored";
            this.checkBox_ExportColored.Size = new System.Drawing.Size(87, 17);
            this.checkBox_ExportColored.TabIndex = 7;
            this.checkBox_ExportColored.Text = "Colored Chat";
            this.checkBox_ExportColored.UseVisualStyleBackColor = true;
            this.checkBox_ExportColored.MouseHover += new System.EventHandler(this.checkBox_ExportColored_MouseHover);
            //
            // checkBox_ExportSplit
            //
            this.checkBox_ExportSplit.AutoSize = true;
            this.checkBox_ExportSplit.Location = new System.Drawing.Point(35, 158);
            this.checkBox_ExportSplit.Name = "checkBox_ExportSplit";
            this.checkBox_ExportSplit.Size = new System.Drawing.Size(71, 17);
            this.checkBox_ExportSplit.TabIndex = 8;
            this.checkBox_ExportSplit.Text = "Split Chat";
            this.checkBox_ExportSplit.UseVisualStyleBackColor = true;
            this.checkBox_ExportSplit.MouseHover += new System.EventHandler(this.checkBox_ExportSplit_MouseHover);
            //
            // checkBox_ExportHtml
            //
            this.checkBox_ExportHtml.AutoSize = true;
            this.checkBox_ExportHtml.Checked = true;
            this.checkBox_ExportHtml.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ExportHtml.Location = new System.Drawing.Point(15, 174);
            this.checkBox_ExportHtml.Name = "checkBox_ExportHtml";
            this.checkBox_ExportHtml.Size = new System.Drawing.Size(151, 17);
            this.checkBox_ExportHtml.TabIndex = 9;
            this.checkBox_ExportHtml.Text = "Export results to TSW html";
            this.checkBox_ExportHtml.UseVisualStyleBackColor = true;
            this.checkBox_ExportHtml.MouseHover += new System.EventHandler(this.checkBox_ExportHtml_MouseHover);
            //
            // checkBox_ExportAllies
            //
            this.checkBox_ExportAllies.AutoSize = true;
            this.checkBox_ExportAllies.Checked = true;
            this.checkBox_ExportAllies.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ExportAllies.Location = new System.Drawing.Point(15, 190);
            this.checkBox_ExportAllies.Name = "checkBox_ExportAllies";
            this.checkBox_ExportAllies.Size = new System.Drawing.Size(174, 17);
            this.checkBox_ExportAllies.TabIndex = 10;
            this.checkBox_ExportAllies.Text = "Export results for raid/team only";
            this.checkBox_ExportAllies.UseVisualStyleBackColor = true;
            this.checkBox_ExportAllies.MouseHover += new System.EventHandler(this.checkBox_ExportAllies_MouseHover);
            //
            // checkBox_DontExportShortEnc
            //
            this.checkBox_DontExportShortEnc.AutoSize = true;
            this.checkBox_DontExportShortEnc.Checked = true;
            this.checkBox_DontExportShortEnc.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_DontExportShortEnc.Location = new System.Drawing.Point(15, 206);
            this.checkBox_DontExportShortEnc.Name = "checkBox_DontExportShortEnc";
            this.checkBox_DontExportShortEnc.Size = new System.Drawing.Size(192, 17);
            this.checkBox_DontExportShortEnc.TabIndex = 11;
            this.checkBox_DontExportShortEnc.Text = "Dont export sub 10 sec encounters";
            this.checkBox_DontExportShortEnc.UseVisualStyleBackColor = true;
            this.checkBox_DontExportShortEnc.MouseHover += new System.EventHandler(this.checkBox_DontExportShortEnc_MouseHover);
            //
            // checkBox_EnableTSWAddon
            //
            this.checkBox_EnableTSWAddon.AutoSize = true;
            this.checkBox_EnableTSWAddon.Location = new System.Drawing.Point(15, 222);
            this.checkBox_EnableTSWAddon.Name = "checkBox_EnableTSWAddon";
            this.checkBox_EnableTSWAddon.Size = new System.Drawing.Size(183, 17);
            this.checkBox_EnableTSWAddon.TabIndex = 12;
            this.checkBox_EnableTSWAddon.Text = "Enable TSWACT Addon features";
            this.checkBox_EnableTSWAddon.UseVisualStyleBackColor = true;
            this.checkBox_EnableTSWAddon.CheckedChanged += new System.EventHandler(this.checkBox_EnableTSWAddon_CheckedChanged);
            this.checkBox_EnableTSWAddon.MouseHover += new System.EventHandler(this.checkBox_EnableTSWAddon_MouseHover);
            //
            // checkBox_AutofixDBConf
            //
            this.checkBox_AutofixDBConf.AutoSize = true;
            this.checkBox_AutofixDBConf.Location = new System.Drawing.Point(35, 238);
            this.checkBox_AutofixDBConf.Name = "checkBox_AutofixDBConf";
            this.checkBox_AutofixDBConf.Size = new System.Drawing.Size(148, 17);
            this.checkBox_AutofixDBConf.TabIndex = 13;
            this.checkBox_AutofixDBConf.Text = "Auto fix dbDebug.conf file";
            this.checkBox_AutofixDBConf.UseVisualStyleBackColor = true;
            this.checkBox_AutofixDBConf.MouseHover += new System.EventHandler(this.checkBox_AutofixDBConf_MouseHover);
            //
            // checkBox_HideUnknown
            //
            this.checkBox_HideUnknown.AutoSize = true;
            this.checkBox_HideUnknown.Location = new System.Drawing.Point(15, 254);
            this.checkBox_HideUnknown.Name = "checkBox_HideUnknown";
            this.checkBox_HideUnknown.Size = new System.Drawing.Size(180, 17);
            this.checkBox_HideUnknown.TabIndex = 14;
            this.checkBox_HideUnknown.Text = "Hide the TSW_Unknown entries";
            this.checkBox_HideUnknown.UseVisualStyleBackColor = true;
            this.checkBox_HideUnknown.MouseHover += new System.EventHandler(this.checkBox_HideUnknown_MouseHover);
            //
            // checkBox_SelfDamage
            //
            this.checkBox_SelfDamage.AutoSize = true;
            this.checkBox_SelfDamage.Location = new System.Drawing.Point(15, 270);
            this.checkBox_SelfDamage.Name = "checkBox_SelfDamage";
            this.checkBox_SelfDamage.Size = new System.Drawing.Size(113, 17);
            this.checkBox_SelfDamage.TabIndex = 15;
            this.checkBox_SelfDamage.Text = "Show self-damage";
            this.checkBox_SelfDamage.UseVisualStyleBackColor = true;
            this.checkBox_SelfDamage.MouseHover += new System.EventHandler(this.checkBox_SelfDamage_MouseHover);
            //
            // checkBox_SelfPlayerDamage
            //
            this.checkBox_SelfPlayerDamage.AutoSize = true;
            this.checkBox_SelfPlayerDamage.Location = new System.Drawing.Point(35, 286);
            this.checkBox_SelfPlayerDamage.Name = "checkBox_SelfPlayerDamage";
            this.checkBox_SelfPlayerDamage.Size = new System.Drawing.Size(99, 17);
            this.checkBox_SelfPlayerDamage.TabIndex = 16;
            this.checkBox_SelfPlayerDamage.Text = "Show by Player";
            this.checkBox_SelfPlayerDamage.UseVisualStyleBackColor = true;
            this.checkBox_SelfPlayerDamage.MouseHover += new System.EventHandler(this.checkBox_SelfPlayerDamage_MouseHover);
            //
            // checkBox_ReduceAegis
            //
            this.checkBox_ReduceAegis.AutoSize = true;
            this.checkBox_ReduceAegis.Location = new System.Drawing.Point(15, 302);
            this.checkBox_ReduceAegis.Name = "checkBox_ReduceAegis";
            this.checkBox_ReduceAegis.Size = new System.Drawing.Size(142, 17);
            this.checkBox_ReduceAegis.TabIndex = 17;
            this.checkBox_ReduceAegis.Text = "Reduce Aegis on names";
            this.checkBox_ReduceAegis.UseVisualStyleBackColor = true;
            this.checkBox_ReduceAegis.MouseHover += new System.EventHandler(this.checkBox_ReduceAegis_MouseHover);
            //
            // labelFilterinig
            //
            this.labelFilterinig.AutoSize = true;
            this.labelFilterinig.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFilterinig.Location = new System.Drawing.Point(12, 355);
            this.labelFilterinig.Name = "labelFilterinig";
            this.labelFilterinig.Size = new System.Drawing.Size(43, 13);
            this.labelFilterinig.TabIndex = 18;
            this.labelFilterinig.Text = "Filtering";
            //
            // checkBox_Filter
            //
            this.checkBox_Filter.AutoSize = true;
            this.checkBox_Filter.Location = new System.Drawing.Point(15, 373);
            this.checkBox_Filter.Name = "checkBox_Filter";
            this.checkBox_Filter.Size = new System.Drawing.Size(143, 17);
            this.checkBox_Filter.TabIndex = 19;
            this.checkBox_Filter.Text = "Enable character filtering";
            this.checkBox_Filter.UseVisualStyleBackColor = true;
            this.checkBox_Filter.CheckedChanged += new System.EventHandler(this.checkBox_Filter_CheckedChanged);
            this.checkBox_Filter.MouseHover += new System.EventHandler(this.checkBox_Filter_MouseHover);
            //
            // labelFilteringName
            //
            this.labelFilteringName.AutoSize = true;
            this.labelFilteringName.Location = new System.Drawing.Point(12, 404);
            this.labelFilteringName.Name = "labelFilteringName";
            this.labelFilteringName.Size = new System.Drawing.Size(35, 13);
            this.labelFilteringName.TabIndex = 20;
            this.labelFilteringName.Text = "Name";
            //
            // textBox_FilterName
            //
            this.textBox_FilterName.Enabled = false;
            this.textBox_FilterName.Location = new System.Drawing.Point(53, 401);
            this.textBox_FilterName.MaxLength = 32;
            this.textBox_FilterName.Name = "textBox_FilterName";
            this.textBox_FilterName.Size = new System.Drawing.Size(149, 20);
            this.textBox_FilterName.TabIndex = 21;
            this.textBox_FilterName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_FilterName_KeyDown);
            this.textBox_FilterName.MouseHover += new System.EventHandler(this.textBox_FilterName_MouseHover);
            //
            // button_AddCharacter
            //
            this.button_AddCharacter.Enabled = false;
            this.button_AddCharacter.Location = new System.Drawing.Point(15, 427);
            this.button_AddCharacter.Name = "button_AddCharacter";
            this.button_AddCharacter.Size = new System.Drawing.Size(83, 26);
            this.button_AddCharacter.TabIndex = 22;
            this.button_AddCharacter.Text = "Add";
            this.button_AddCharacter.UseVisualStyleBackColor = true;
            this.button_AddCharacter.Click += new System.EventHandler(this.button_AddCharacter_Click);
            this.button_AddCharacter.MouseHover += new System.EventHandler(this.button_AddCharacter_MouseHover);
            //
            // button_DeleteCharacter
            //
            this.button_DeleteCharacter.Enabled = false;
            this.button_DeleteCharacter.Location = new System.Drawing.Point(119, 427);
            this.button_DeleteCharacter.Name = "button_DeleteCharacter";
            this.button_DeleteCharacter.Size = new System.Drawing.Size(83, 26);
            this.button_DeleteCharacter.TabIndex = 23;
            this.button_DeleteCharacter.Text = "Delete";
            this.button_DeleteCharacter.UseVisualStyleBackColor = true;
            this.button_DeleteCharacter.Click += new System.EventHandler(this.button_DeleteCharacter_Click);
            this.button_DeleteCharacter.MouseHover += new System.EventHandler(this.button_DeleteCharacter_MouseHover);
            //
            // checkBox_filterExclude
            //
            this.checkBox_filterExclude.AutoSize = true;
            this.checkBox_filterExclude.Location = new System.Drawing.Point(15, 469);
            this.checkBox_filterExclude.Name = "checkBox_filterExclude";
            this.checkBox_filterExclude.Size = new System.Drawing.Size(107, 17);
            this.checkBox_filterExclude.TabIndex = 24;
            this.checkBox_filterExclude.Text = "Exclusive filtering";
            this.checkBox_filterExclude.UseVisualStyleBackColor = true;
            this.checkBox_filterExclude.MouseHover += new System.EventHandler(this.checkBox_filterExclude_MouseHover);
            //
            // checkBox_filterScript
            //
            this.checkBox_filterScript.AutoSize = true;
            this.checkBox_filterScript.Location = new System.Drawing.Point(15, 492);
            this.checkBox_filterScript.Name = "checkBox_filterScript";
            this.checkBox_filterScript.Size = new System.Drawing.Size(131, 17);
            this.checkBox_filterScript.TabIndex = 25;
            this.checkBox_filterScript.Text = "Filter Generated Script";
            this.checkBox_filterScript.UseVisualStyleBackColor = true;
            this.checkBox_filterScript.MouseHover += new System.EventHandler(this.checkBox_filterScript_MouseHover);
            //
            // labelChatExport
            //
            this.labelChatExport.AutoSize = true;
            this.labelChatExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelChatExport.Location = new System.Drawing.Point(275, 92);
            this.labelChatExport.Name = "labelChatExport";
            this.labelChatExport.Size = new System.Drawing.Size(92, 13);
            this.labelChatExport.TabIndex = 26;
            this.labelChatExport.Text = "Chat Export Fields";
            //
            // checkedListBox_ExportFields
            //
            this.checkedListBox_ExportFields.Enabled = false;
            this.checkedListBox_ExportFields.FormattingEnabled = true;
            this.checkedListBox_ExportFields.Items.AddRange(new object[] {
            "DPS",
            "Crit%",
            "Pen%",
            "Glance%",
            "Block%",
            "HPS",
            "HealCrit%",
            "Evade%",
            "AegisMismatch%",
            "TakenDamage",
            "TakenCrit%",
            "TakenPen%",
            "TakenGlance%",
            "TakenBlock%",
            "TakenEvade%"});
            this.checkedListBox_ExportFields.Location = new System.Drawing.Point(222, 110);
            this.checkedListBox_ExportFields.Name = "checkedListBox_ExportFields";
            this.checkedListBox_ExportFields.Size = new System.Drawing.Size(227, 139);
            this.checkedListBox_ExportFields.TabIndex = 27;
            //
            // labelFilterChars
            //
            this.labelFilterChars.AutoSize = true;
            this.labelFilterChars.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelFilterChars.Location = new System.Drawing.Point(275, 261);
            this.labelFilterChars.Name = "labelFilterChars";
            this.labelFilterChars.Size = new System.Drawing.Size(83, 13);
            this.labelFilterChars.TabIndex = 28;
            this.labelFilterChars.Text = "Filter Characters";
            //
            // checkedListBox_Filters
            //
            this.checkedListBox_Filters.Enabled = false;
            this.checkedListBox_Filters.FormattingEnabled = true;
            this.checkedListBox_Filters.Location = new System.Drawing.Point(222, 279);
            this.checkedListBox_Filters.Name = "checkedListBox_Filters";
            this.checkedListBox_Filters.Size = new System.Drawing.Size(227, 199);
            this.checkedListBox_Filters.TabIndex = 29;
            this.checkedListBox_Filters.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBox_Filters_ItemCheck);
            this.checkedListBox_Filters.MouseHover += new System.EventHandler(this.checkedListBox_Filters_MouseHover);
            //
            // SecretParse
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.checkBox_SelfPlayerDamage);
            this.Controls.Add(this.labelHeader);
            this.Controls.Add(this.labelGeneral);
            this.Controls.Add(this.labelLanguage);
            this.Controls.Add(this.comboBox_Language);
            this.Controls.Add(this.checkBox_AutoCheck);
            this.Controls.Add(this.checkBox_ExportScript);
            this.Controls.Add(this.checkBox_LimitNames);
            this.Controls.Add(this.checkBox_ExportRoundDPS);
            this.Controls.Add(this.checkBox_ExportShowLegend);
            this.Controls.Add(this.checkBox_ExportColored);
            this.Controls.Add(this.checkBox_ExportSplit);
            this.Controls.Add(this.checkBox_ExportHtml);
            this.Controls.Add(this.checkBox_ExportAllies);
            this.Controls.Add(this.checkBox_DontExportShortEnc);
            this.Controls.Add(this.checkBox_EnableTSWAddon);
            this.Controls.Add(this.checkBox_AutofixDBConf);
            this.Controls.Add(this.checkBox_HideUnknown);
            this.Controls.Add(this.checkBox_SelfDamage);
            this.Controls.Add(this.checkBox_ReduceAegis);
            this.Controls.Add(this.labelFilterinig);
            this.Controls.Add(this.checkBox_Filter);
            this.Controls.Add(this.labelFilteringName);
            this.Controls.Add(this.textBox_FilterName);
            this.Controls.Add(this.button_AddCharacter);
            this.Controls.Add(this.button_DeleteCharacter);
            this.Controls.Add(this.checkBox_filterExclude);
            this.Controls.Add(this.checkBox_filterScript);
            this.Controls.Add(this.labelChatExport);
            this.Controls.Add(this.checkedListBox_ExportFields);
            this.Controls.Add(this.labelFilterChars);
            this.Controls.Add(this.checkedListBox_Filters);
            this.Name = "SecretParse";
            this.Size = new System.Drawing.Size(494, 534);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelHeader;
        private System.Windows.Forms.Label labelGeneral;
        private System.Windows.Forms.Label labelLanguage;
        private System.Windows.Forms.ComboBox comboBox_Language;
        private System.Windows.Forms.CheckBox checkBox_AutoCheck;
        private System.Windows.Forms.CheckBox checkBox_ExportScript;
        private System.Windows.Forms.CheckBox checkBox_LimitNames;
        private System.Windows.Forms.CheckBox checkBox_ExportRoundDPS;
        private System.Windows.Forms.CheckBox checkBox_ExportShowLegend;
        private System.Windows.Forms.CheckBox checkBox_ExportColored;
        private System.Windows.Forms.CheckBox checkBox_ExportSplit;
        private System.Windows.Forms.CheckBox checkBox_ExportHtml;
        private System.Windows.Forms.CheckBox checkBox_ExportAllies;
        private System.Windows.Forms.CheckBox checkBox_DontExportShortEnc;
        private System.Windows.Forms.CheckBox checkBox_EnableTSWAddon;
        private System.Windows.Forms.CheckBox checkBox_AutofixDBConf;
        private System.Windows.Forms.CheckBox checkBox_HideUnknown;
        private System.Windows.Forms.CheckBox checkBox_SelfDamage;
        private System.Windows.Forms.CheckBox checkBox_SelfPlayerDamage;
        private System.Windows.Forms.CheckBox checkBox_ReduceAegis;
        private System.Windows.Forms.Label labelFilterinig;
        private System.Windows.Forms.CheckBox checkBox_Filter;
        private System.Windows.Forms.Label labelFilteringName;
        private System.Windows.Forms.TextBox textBox_FilterName;
        private System.Windows.Forms.Button button_AddCharacter;
        private System.Windows.Forms.Button button_DeleteCharacter;
        private System.Windows.Forms.CheckBox checkBox_filterExclude;
        private System.Windows.Forms.CheckBox checkBox_filterScript;
        private System.Windows.Forms.Label labelChatExport;
        private System.Windows.Forms.CheckedListBox checkedListBox_ExportFields;
        private System.Windows.Forms.Label labelFilterChars;
        private System.Windows.Forms.CheckedListBox checkedListBox_Filters;
        #endregion


        public SecretParse()
        {
            InitializeComponent();
        }

        private const string ALL = "All";
        private const string OUT_DAMAGE = "Attack (Out)";
        private const string OUT_HEAL = "Healed (Out)";
        private const string ALL_OUTGOING = "All Outgoing (Ref)";
        private const string INC_DAMAGE = "Incoming Damage";
        private const int CHAT_LIMIT = 2400;
        private Label lblStatus;
        private TreeNode optionsNode = null;
        private string settingsFile = Path.Combine(ActGlobals.oFormActMain.AppDataFolder.FullName, "Secret.config.xml");
        private SettingsSerializer xmlSettings;
        private HashSet<string> filterNames = new HashSet<string>();
        private int healCount = 0;
        private int dpsCount = 0;
        private static int runClientReadThread = 1;
        private object locker = new object();
        private object addonEnabledLocker = new object();
        private bool addonEnabled;
        private string charName = "";
        private string combatEnding = null;
        private DateTime lastCombatStart = DateTime.MinValue;
        private DateTime lastCombatEnd = DateTime.MinValue;
        private DateTime lastEvadeTime = DateTime.MinValue;
        private Dictionary<string, BuffData> buffData = new Dictionary<string, BuffData>();
        private HashSet<string> allies = new HashSet<string>();
        private Dictionary<string, int> exportFieldMap = new Dictionary<string, int>();
        private string previousLine = "";
        private static Regex timeStamp = new Regex(@"^(\[[0-9]{2}:[0-9]{2}:[0-9]{2}\]\s*)+", RegexOptions.Compiled);
        private static Regex fontHtml = new Regex(@"</?font[^>]*>", RegexOptions.Compiled);
        private static string BUFFS = "BUFFS";
        private static string HTML_HEADER = @"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">
<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""en-US"">
<head>
	<meta charset=""utf-8"">
	<script type=""text/javascript"" src=""https://ajax.googleapis.com/ajax/libs/jquery/1.10.0/jquery.min.js""></script>
	<script>
		$(document).ready(function() {
			$(""#content div"").hide();
			$(""#tabs li:first"").attr(""id"",""current"");
			$(""#content div:first"").fadeIn();
			$('#tabs a').click(function(e) {
				e.preventDefault();
				$(""#content div"").hide();
				$(""#tabs li"").attr(""id"","""");
				$(this).parent().attr(""id"",""current"");
				$('#' + $(this).attr('title')).fadeIn();
			});
		});
		function toggle(obj) {
			var tr = document.getElementsByTagName(""tr"");
			for (i=0;i<tr.length;i++) {
				if (tr[i].className == obj) {
					tr[i].style.display = (tr[i].style.display!='none' ? 'none' : '');
				}
			}
		}
	</script>
	<style>
		::-webkit-scrollbar {
			width: 6px;
		}
		::-webkit-scrollbar-track {
			border-radius: 6px;
			-webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0);
		}
		::-webkit-scrollbar-thumb {
			border-radius: 6px;
			background: rgba(34,123,212,.7);
			-webkit-box-shadow: inset 0 0 6px rgba(255,255,255,.3);
		}
		html {
			overflow-y: scroll;
		}
		body {
			font: 76%/160% ""Trebuchet MS"", sans-serif;
			margin: 8px;
		}
		h1 {
			margin: 0;
			padding: 0 15px 0 0;
			font:bold 22px/22px ""Trebuchet MS"", sans-serif;
			color:#f72c17;
			text-align:right;
		}
		#tabs {
			overflow: auto;
			width: 100%;
			list-style: none;
			margin: 0;
			padding: 0;
		}
		#tabs li {
			margin: 0;
			padding: 0;
			float: left;
		}
		#tabs li.title {
			margin: -5px 0 0 0;
			padding: 0 15px 0 0;
			float: right;
			font:bold 16px/35px ""Trebuchet MS"", sans-serif;
			color:#fb8b13;
		}
		#tabs a {
			background: #d0e5f5;
			background: -webkit-linear-gradient(226deg, transparent 10px, #d0e5f5 10px);
			background: -moz-linear-gradient(226deg, transparent 10px, #d0e5f5 10px);
			background: -ms-linear-gradient(226deg, transparent 10px, #d0e5f5 10px);
			background: -o-linear-gradient(226deg, transparent 10px, #d0e5f5 10px);
			background: linear-gradient(226deg, transparent 10px, #d0e5f5 10px);
			-webkit-box-shadow: -4px 0 0 rgba(0, 0, 0, .2);
			-moz-box-shadow: -4px 0 0 rgba(0, 0, 0, .2);
			box-shadow: -4px 0 0 rgba(0, 0, 0, .2);
			font: bold 13px/35px 'Lucida sans', Arial, Helvetica;
			color: #1d5987;
			text-decoration: none;
			float: left;
			height: 35px;
			padding: 0 30px;
		}
		#tabs a:hover {
			background: #4e92d9;
			background: -webkit-linear-gradient(226deg, transparent 10px, #4e92d9 10px);
			background: -moz-linear-gradient(226deg, transparent 10px, #4e92d9 10px);
			background: -ms-linear-gradient(226deg, transparent 10px, #4e92d9 10px);
			background: -o-linear-gradient(226deg, transparent 10px, #4e92d9 10px);
			background: linear-gradient(226deg, transparent 10px, #4e92d9 10px);
			color: #fff;
		}
		#tabs a:focus {
			outline: 0;
		}
		#tabs #current a {
			background: #f6faff;
			background: -webkit-linear-gradient(226deg, transparent 10px, #f6faff 10px);
			background: -moz-linear-gradient(226deg, transparent 10px, #f6faff 10px);
			background: -ms-linear-gradient(226deg, transparent 10px, #f6faff 10px);
			background: -o-linear-gradient(226deg, transparent 10px, #f6faff 10px);
			background: linear-gradient(226deg, transparent 10px, #f6faff 10px);
			color: #d54e0b;
		}
		#content {
			background-color: #f6faff;
			background-image: -webkit-gradient(linear, left top, left bottom, from(#f6faff), to(#bedbff));
			background-image: -webkit-linear-gradient(top, #f6faff, #bedbff);
			background-image: -moz-linear-gradient(top, #f6faff, #bedbff);
			background-image: -ms-linear-gradient(top, #f6faff, #bedbff);
			background-image: -o-linear-gradient(top, #f6faff, #bedbff);
			background-image: linear-gradient(top, #f6faff, #bedbff);
			-webkit-border-radius: 0 6px 6px 6px;
			-moz-border-radius: 0 6px 6px 6px;
			border-radius: 0 6px 6px 6px;
			padding: 6px;
		}
		#content table {
			width:100%;
			border-collapse:collapse;
		}
		#content th {
			background:#e2f0fd;
			border: 1px solid #dae8f5;
			text-align:center;
			font:bold 13px/16px ""Trebuchet MS"", sans-serif;
			color:#0a2e74;
			height:40px;
		}
		#content td {
			border:1px solid #dae8f5;
			text-align:right;
			color:#678197;
			padding:0 5px;
		}
		#content tfoot td {
			background:#e2f0fd;
			border: 1px solid #dae8f5;
			text-align:center;
			font:italic 12px/22px ""Trebuchet MS"", sans-serif;
			color:#636363;
		}
		#content th.heal, #content td.heal {
			border-left:5px solid #dae8f5;
		}
		#content td.col1 {
			text-align:left;
			padding:3px 5px;
		}
		#content td strong {
			font:bold 13px/28px ""Trebuchet MS"", sans-serif;
			color:#2937FF;
		}
		#content td b {
			font-weight:normal;
			color:#434DDE;
			padding:0 10px;
		}
		#content tr.even, tr.buffeven {
			background:#f0f7fe;
		}
		#content tr.odd, tr.buffodd {
			background:#f7fbff;
		}
		#content tr.even:hover, #content tr.odd:hover {
			background:#FCFFB0;
			cursor:pointer;
		}
		#content tr.buffeven:hover, #content tr.buffodd:hover {
			background:#FCFFB0;
		}
";
        private static string HTML_TITLE = @"
	</style>
</head>
<body>
	<h1>{0}</h1>
	<ul id=""tabs"">
		<li><a href="""" title=""tab1"">Group Damage</a></li>
		<li><a href="""" title=""tab2"">My Buffs</a></li>
		<li><a href="""" title=""tab3"">Big Numbers</a></li>
		<li class=""title"">total: {3} dmg, {1} dps in {2}</li>
	</ul>";
        private static string HTML_TABLE_DAMAGE = @"
	<div id=""content"">
		<div id=""tab1"">
			<table>
				<thead>
					<tr>
						<th>Name - (combat time)</th>
						<th width=""8%"">Damage</th>
						<th width=""6%"">DPS</th>
						<th width=""6%"">Pen</th>
						<th width=""6%"">Crit</th>
						<th width=""6%"">Glnc</th>
						<th width=""6%"">Block</th>
						<th width=""6%"">Evade</th>
						<th width=""3%"">Aegis damage</th>
						<th width=""3%"">Aegis DPS</th>
						<th width=""3%"">Aegis mismatch</th>
						<th width=""8%"" class=""heal"">Heal</th>
						<th width=""6%"">HPS</th>
						<th width=""6%"">HCrit</th>
						<th width=""3%"">Aegis heal</th>
						<th width=""3%"">Aegis HPS</th>
					</tr>
				</thead>
				<tfoot>
					<tr>
						<td colspan=""16"">Click on a row to show the detailed skill breakdown</td>
					</tr>
				</tfoot>
				<tbody>
					";
        private static string HTML_TABLE_BUFFS = @"
				</tbody>
			</table>
		</div>
		<div id=""tab2"">
			<table>
				<thead>
					<tr>
						<th>Name</th>
						<th width=""18%"">Duration</th>
						<th width=""18%"">Percent</th>
					</tr>
				</thead>
				<tbody>
					";
        private static string HTML_TABLE_MAX = @"
				</tbody>
			</table>
		</div>
		<div id=""tab3"">
			<table>
				<thead>
					<tr>
						<th>Name</th>
						<th width=""25%"">Max Hit</th>
						<th width=""25%"">Max Heal</th>
						<th width=""25%"">Max Incoming</th>
					</tr>
				</thead>
				<tbody>
					";
        private static string HTML_CLOSE = @"
				</tbody>
			</table>
		</div>
	</div>
</body>
</html>";

        public void InitPlugin(TabPage pluginScreenSpace, Label pluginStatusText)
        {
            lblStatus = pluginStatusText;

            // Push the option screen into the option tab
            int dcIndex = -1;
            for (int i = 0; i < ActGlobals.oFormActMain.OptionsTreeView.Nodes.Count; i++)
            {
                if (ActGlobals.oFormActMain.OptionsTreeView.Nodes[i].Text == "Secret Parsing")
                    dcIndex = i;
            }
            if (dcIndex != -1)
            {
                optionsNode = ActGlobals.oFormActMain.OptionsTreeView.Nodes[dcIndex].Nodes.Add("General");
                ActGlobals.oFormActMain.OptionsControlSets.Add(@"Secret Parsing\General", new List<Control> { this });
                Label lblConfig = new Label();
                lblConfig.AutoSize = true;
                lblConfig.Text = "Find the applicable options in the Options tab, Secret Parsing section.";
                pluginScreenSpace.Controls.Add(lblConfig);
            }
            else
            {
                ActGlobals.oFormActMain.OptionsTreeView.Nodes.Add("Secret Parsing");
                dcIndex = ActGlobals.oFormActMain.OptionsTreeView.Nodes.Count - 1;
                optionsNode = ActGlobals.oFormActMain.OptionsTreeView.Nodes[dcIndex].Nodes.Add("General");
                ActGlobals.oFormActMain.OptionsControlSets.Add(@"Secret Parsing\General", new List<Control> { this });
                Label lblConfig = new Label();
                lblConfig.AutoSize = true;
                lblConfig.Text = "Find the applicable options in the Options tab, Secret Parsing section.";
                pluginScreenSpace.Controls.Add(lblConfig);
            }
            ActGlobals.oFormActMain.OptionsTreeView.Nodes[dcIndex].Expand();
            ActGlobals.oFormActMain.SetOptionsHelpText("testing");

            // Add the export item to the treeview menu
            var exportButton = new ToolStripMenuItem();
            exportButton.Name = "ACT_export_button";
            exportButton.Text = "Export to TSW chat script";
            exportButton.Click += new System.EventHandler(ExportEncounterToScript);
            int sepIndx = 0;
            for (sepIndx = 0; sepIndx < ActGlobals.oFormActMain.MainTreeView.ContextMenuStrip.Items.Count; ++sepIndx)
            {
                if (ActGlobals.oFormActMain.MainTreeView.ContextMenuStrip.Items[sepIndx] is ToolStripSeparator)
                {
                    break;
                }
            }
            ActGlobals.oFormActMain.MainTreeView.ContextMenuStrip.Items.Insert(sepIndx, exportButton);
            ActGlobals.oFormActMain.MainTreeView.ContextMenuStrip.Opening += new CancelEventHandler((a, b) =>
            {
                exportButton.Enabled = IsEncounterSelected();
            });

            xmlSettings = new SettingsSerializer(this);
            string FileName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Secret\\recents.cfg";
            FileInfo RecFile = new FileInfo(FileName);

            SetExportFieldDefaults();

            LoadSettings();

            SetExportFieldStatus();

            checkBox_EnableTSWAddon_CheckedChanged(null, null);

            SetupSecretEnglishEnvironment();

            //ActGlobals.oFormActMain.SetParserToNull(); // Set the input log file language to (None)
            ActGlobals.oFormActMain.LogPathHasCharName = false;
            ActGlobals.oFormActMain.LogFileFilter = "CombatLog*.txt";
            ActGlobals.oFormActMain.ResetCheckLogs();

            ActGlobals.oFormActMain.TimeStampLen = 8; // Size of timestamp section in the logfile
            ActGlobals.oFormActMain.GetDateTimeFromLog = new FormActMain.DateTimeLogParser(ParseDateTime); // Replace internal Eq2 date parser

            ActGlobals.oFormActMain.BeforeLogLineRead += new LogLineEventDelegate(oFormActMain_BeforeLogLineRead); // Interupt the EQ2 parse, and replace it with Secret.
            ActGlobals.oFormActMain.OnCombatEnd += new CombatToggleEventDelegate(oFormActMain_OnCombatEnd);
            ActGlobals.oFormActMain.OnCombatStart += new CombatToggleEventDelegate(oFormActMain_OnCombatStart);

            // Setup up update checks
            ActGlobals.oFormActMain.UpdateCheckClicked += new FormActMain.NullDelegate(SecretCheckUpdate);
            if (ActGlobals.oFormActMain.GetAutomaticUpdatesAllowed() == true)
            {
                new Thread(new ThreadStart(SecretCheckUpdate)).Start();
            }

            FixDBConfFile();

            Task.Factory.StartNew(() => { CheckTSWACTEnabled(this); });

            if ((lblStatus.Text != "Secret plugin unloaded") && (lblStatus.Text != "No Status"))
            {
                lblStatus.Text = lblStatus.Text + "\nSecret plugin loaded";
            }
            else
            {
                lblStatus.Text = "Secret plugin loaded";
            }
        }

        private bool IsExportFieldSet(string name)
        {
            string lowName = name.ToLower();
            if (exportFieldMap.ContainsKey(lowName))
            {
                return checkedListBox_ExportFields.GetItemChecked(exportFieldMap[lowName]);
            }

            return false;
        }

        private void SetExportFieldDefaults()
        {
            checkedListBox_ExportFields.SetItemChecked(0, true);  // DPS
            checkedListBox_ExportFields.SetItemChecked(1, true);  // Crit%
            checkedListBox_ExportFields.SetItemChecked(2, true);  // Pen%
            checkedListBox_ExportFields.SetItemChecked(3, true);  // Glance%
            checkedListBox_ExportFields.SetItemChecked(4, false); // Block%
            checkedListBox_ExportFields.SetItemChecked(5, true);  // HPS
            checkedListBox_ExportFields.SetItemChecked(6, true);  // HealCrit%
            checkedListBox_ExportFields.SetItemChecked(7, true);  // Evade%
            checkedListBox_ExportFields.SetItemChecked(8, false); // AegisMismatch%
            checkedListBox_ExportFields.SetItemChecked(9, true);  // TakenDamage
            checkedListBox_ExportFields.SetItemChecked(10, true); // TakenCrit%
            checkedListBox_ExportFields.SetItemChecked(11, true); // TakenPen%
            checkedListBox_ExportFields.SetItemChecked(12, true); // TakenGlance%
            checkedListBox_ExportFields.SetItemChecked(13, true); // TakenBlock%
            checkedListBox_ExportFields.SetItemChecked(14, true); // TakenEvade%

            exportFieldMap.Add("dps", 0);
            exportFieldMap.Add("crit%", 1);
            exportFieldMap.Add("pen%", 2);
            exportFieldMap.Add("glance%", 3);
            exportFieldMap.Add("block%", 4);
            exportFieldMap.Add("hps", 5);
            exportFieldMap.Add("healcrit%", 6);
            exportFieldMap.Add("evade%", 7);
            exportFieldMap.Add("aegismismatch%", 8);
            exportFieldMap.Add("takendamage", 9);
            exportFieldMap.Add("takencrit%", 10);
            exportFieldMap.Add("takenpen%", 11);
            exportFieldMap.Add("takenglance%", 12);
            exportFieldMap.Add("takenblock%", 13);
            exportFieldMap.Add("takenevade%", 14);
        }


        public void DeInitPlugin()
        {
            Interlocked.Decrement(ref runClientReadThread);

            ActGlobals.oFormActMain.BeforeLogLineRead -= oFormActMain_BeforeLogLineRead;
            ActGlobals.oFormActMain.OnCombatEnd -= oFormActMain_OnCombatEnd;
            ActGlobals.oFormActMain.UpdateCheckClicked -= SecretCheckUpdate;

            if (optionsNode != null)
            {
                optionsNode.Remove();
                ActGlobals.oFormActMain.OptionsControlSets.Remove(@"Secret Parsing\General");
            }

            for (int i = 0; i < ActGlobals.oFormActMain.OptionsTreeView.Nodes.Count; i++)
            {
                if (ActGlobals.oFormActMain.OptionsTreeView.Nodes[i].Text == "Secret Parsing")
                    ActGlobals.oFormActMain.OptionsTreeView.Nodes[i].Remove();
            }

            for (int sepIndx = 0; sepIndx < ActGlobals.oFormActMain.MainTreeView.ContextMenuStrip.Items.Count; ++sepIndx)
            {
                if (ActGlobals.oFormActMain.MainTreeView.ContextMenuStrip.Items[sepIndx] is ToolStripMenuItem)
                {
                    ToolStripMenuItem tempItem = (ToolStripMenuItem)ActGlobals.oFormActMain.MainTreeView.ContextMenuStrip.Items[sepIndx];
                    if ("ACT_export_button".Equals(tempItem.Name))
                    {
                        ActGlobals.oFormActMain.MainTreeView.ContextMenuStrip.Items.RemoveAt(sepIndx);
                        break;
                    }
                }
            }

            SaveSettings();

            FixDBConfFile();

            lblStatus.Text = "Secret plugin unloaded";
        }

        private long GetHealData(EncounterData Data, string specialName)
        {
            long total = 0;
            List<CombatantData> allys = Data.GetAllies();

            foreach (var ally in allys)
            {
                if (ally.Items.ContainsKey(OUT_HEAL) && ally.Items[OUT_HEAL].Items.ContainsKey(ALL))
                {
                    foreach (var swing in ally.Items[OUT_HEAL].Items[ALL].Items)
                    {
                        total += swing.Damage;
                    }

                }
            }

            return total;
        }

        private long GetSpecialHitData(EncounterData Data, string specialName)
        {
            long total = 0;
            List<CombatantData> allys = Data.GetAllies();

            foreach (var ally in allys)
            {
                if (ally.Items.ContainsKey(OUT_DAMAGE) && ally.Items[OUT_DAMAGE].Items.ContainsKey(ALL))
                {
                    foreach (var swing in ally.Items[OUT_DAMAGE].Items[ALL].Items)
                    {
                        if (swing.Special.Contains(specialName))
                        {
                            total += swing.Damage;
                        }
                    }

                }
            }

            return total;
        }

        private long GetSpecialHealData(EncounterData Data, string specialName)
        {
            long total = 0;
            List<CombatantData> allys = Data.GetAllies();

            foreach (var ally in allys)
            {
                if (ally.Items.ContainsKey(OUT_HEAL) && ally.Items[OUT_HEAL].Items.ContainsKey(ALL))
                {
                    foreach (var swing in ally.Items[OUT_HEAL].Items[ALL].Items)
                    {
                        if (swing.Special.Contains(specialName))
                        {
                            total += swing.Damage;
                        }
                    }

                }
            }

            return total;
        }

        private long GetSpecialHitData(CombatantData Data, string specialName)
        {
            long total = 0;

            if (Data.Items.ContainsKey(OUT_DAMAGE) && Data.Items[OUT_DAMAGE].Items.ContainsKey(ALL))
            {
                foreach (var swing in Data.Items[OUT_DAMAGE].Items[ALL].Items)
                {
                    if (swing.Special.Contains(specialName))
                    {
                        total += swing.Damage;
                    }
                }
            }

            return total;
        }

        private long GetSpecialHealData(CombatantData Data, string specialName)
        {
            long total = 0;

            if (Data.Items.ContainsKey(OUT_HEAL) && Data.Items[OUT_HEAL].Items.ContainsKey(ALL))
            {
                foreach (var swing in Data.Items[OUT_HEAL].Items[ALL].Items)
                {
                    if (swing.Special.Contains(specialName))
                    {
                        total += swing.Damage;
                    }
                }
            }

            return total;
        }

        private long GetSpecialIncData(CombatantData Data, string specialName)
        {
            long total = 0;

            if (Data.Items.ContainsKey(INC_DAMAGE) && Data.Items[INC_DAMAGE].Items.ContainsKey(ALL))
            {
                foreach (var swing in Data.Items[INC_DAMAGE].Items[ALL].Items)
                {
                    if (swing.Special.Contains(specialName))
                    {
                        total += swing.Damage;
                    }
                }
            }

            return total;
        }

        private long GetSpecialHitData(DamageTypeData Data, string specialName)
        {
            long total = 0;

            if (Data.Items.ContainsKey(ALL))
            {
                foreach (var swing in Data.Items[ALL].Items)
                {
                    if (swing.Special.Contains(specialName))
                    {
                        total += swing.Damage;
                    }
                }
            }

            return total;
        }

        private long GetSpecialHitData(AttackType Data, string specialName)
        {
            long total = 0;

            foreach (var swing in Data.Items)
            {
                if (swing.Special.Contains(specialName))
                {
                    total += swing.Damage;
                }
            }

            return total;
        }

        private int GetSpecialHitCount(DamageTypeData Data, string specialName)
        {
            int count = 0;

            if (Data.Items.ContainsKey(ALL))
            {
                foreach (var swing in Data.Items[ALL].Items)
                {
                    if (swing.Special.Contains(specialName))
                    {
                        ++count;
                    }
                }
            }

            return count;
        }

        private int GetSpecialHitCount(AttackType Data, string specialName)
        {
            int count = 0;

            foreach (var swing in Data.Items)
            {
                if (swing.Special.Contains(specialName))
                {
                    ++count;
                }
            }

            return count;
        }

        private double GetSpecialHitPerc(DamageTypeData Data, string specialName)
        {
            if ((Data.Hits + Data.Misses) < 1)
                return 0;

            double specials = GetSpecialHitCount(Data, specialName);
            specials /= (Data.Hits + Data.Misses);
            return specials * 100.0;
        }

        private int GetSpecialHitMiss(DamageTypeData Data, string specialName) {
            int count = 0;

            if (Data.Items.ContainsKey(ALL)) {
                foreach (var swing in Data.Items[ALL].Items)
                {
                    if (swing.Special.Contains(specialName) && (swing.Damage == 0))
                    {
                        ++count;
                    }
                }
            }

            return count;
        }

        private double GetSpecialHitMissPerc(DamageTypeData Data, string specialName) {
            int hits = GetSpecialHitCount(Data, specialName);
            if (hits < 1)
                return 0;

            double specials = GetSpecialHitMiss(Data, specialName);
            specials /= hits;
            return specials * 100.0;
        }

        private double GetSpecialHitPerc(AttackType Data, string specialName)
        {
            if ((Data.Hits + Data.Misses) < 1)
                return 0;

            double specials = GetSpecialHitCount(Data, specialName);
            specials /= (Data.Hits + Data.Misses);
            return specials * 100.0;
        }

        private int GetSpecialHitMiss(AttackType Data, string specialName)
        {
            int count = 0;

            foreach (var swing in Data.Items)
            {
                if (swing.Special.Contains(specialName) && (swing.Damage == 0))
                {
                    ++count;
                }
            }

            return count;
        }

        private double GetSpecialHitMissPerc(AttackType Data, string specialName)
        {
            int hits = GetSpecialHitCount(Data, specialName);
            if (hits < 1)
              return 0;

            double specials = GetSpecialHitMiss(Data, specialName);
            specials /= hits;
            return specials * 100.0;
        }

        private double CalcDataPS(long Dmg, double Seconds)
        {
            double result = 0.0;
            if (Seconds != 0)
            {
                result = (float)Dmg / Seconds;
            }
            return result;
        }

        private string CombatantFormatSwitch(CombatantData Data, string VarName, CultureInfo usCulture)
        {
            if ((!Data.Items.ContainsKey(OUT_DAMAGE)) || (!Data.Items[OUT_DAMAGE].Items.ContainsKey(ALL)))
            {
                return "0%";
            }

            switch (VarName)
            {
                case "Glance%":
                    return GetSpecialHitPerc(Data.Items[OUT_DAMAGE].Items[ALL], SecretLanguage.Glancing).ToString("0'%", usCulture);
                case "Penetration%":
                    return GetSpecialHitPerc(Data.Items[OUT_DAMAGE].Items[ALL], SecretLanguage.Penetrated).ToString("0'%", usCulture);
                case "Blocked%":
                    return GetSpecialHitPerc(Data.Items[OUT_DAMAGE].Items[ALL], SecretLanguage.Blocked).ToString("0'%", usCulture);
                case "AegisMismatch%":
                    return GetSpecialHitMissPerc(Data.Items[OUT_DAMAGE].Items[ALL], SecretLanguage.Aegis).ToString("0'%", usCulture);
                default:
                    return VarName;
            }
        }

        private string CombatantFormatIncSwitch(CombatantData Data, string VarName, CultureInfo usCulture)
        {
            if ((!Data.Items.ContainsKey(INC_DAMAGE)) || (!Data.Items[INC_DAMAGE].Items.ContainsKey(ALL)))
            {
                return "0%";
            }

            switch (VarName)
            {
                case "takencrit%":
                    return Data.Items[INC_DAMAGE].Items[ALL].CritPerc.ToString("0'%", usCulture);
                case "takenpen%":
                    return GetSpecialHitPerc(Data.Items[INC_DAMAGE].Items[ALL],SecretLanguage.Penetrated).ToString("0'%", usCulture);
                case "takenglance%":
                    return GetSpecialHitPerc(Data.Items[INC_DAMAGE].Items[ALL],SecretLanguage.Glancing).ToString("0'%", usCulture);
                case "takenblock%":
                    return GetSpecialHitPerc(Data.Items[INC_DAMAGE].Items[ALL],SecretLanguage.Blocked).ToString("0'%", usCulture);
                case "takenevade%":
                    double missperc = 0.0;
                    if ((Data.Items[INC_DAMAGE].Items[ALL].Hits + Data.Items[INC_DAMAGE].Items[ALL].Misses) > 0)
                    {
                        missperc = 100.0 * Data.Items[INC_DAMAGE].Items[ALL].Misses / (Data.Items[INC_DAMAGE].Items[ALL].Hits + Data.Items[INC_DAMAGE].Items[ALL].Misses);
                    }
                    return missperc.ToString("0'%", usCulture);
                default:
                    return VarName;
             }
        }

        private void SetupSecretEnglishEnvironment()
        {
            CultureInfo usCulture = new CultureInfo("en-US");	// This is for SQL syntax; do not change

            CombatantData.OutgoingDamageTypeDataObjects = new Dictionary<string, CombatantData.DamageTypeDef>
            {
              {OUT_DAMAGE, new CombatantData.DamageTypeDef(OUT_DAMAGE, -1, Color.DarkGoldenrod)},
              {"Outgoing Damage", new CombatantData.DamageTypeDef("Outgoing Damage", 0, Color.Orange)},
              {OUT_HEAL, new CombatantData.DamageTypeDef(OUT_HEAL, 1, Color.Blue)},
              {"Outgoing Heal (Out)", new CombatantData.DamageTypeDef("Outgoing Heal (Out)", 1, Color.Blue)},
              {ALL_OUTGOING, new CombatantData.DamageTypeDef(ALL_OUTGOING, 0, Color.Black)}
            };
            CombatantData.IncomingDamageTypeDataObjects = new Dictionary<string, CombatantData.DamageTypeDef>
            {
              {INC_DAMAGE, new CombatantData.DamageTypeDef(INC_DAMAGE, -1, Color.Red)},
              {"Healed (Inc)",new CombatantData.DamageTypeDef("Healed (Inc)", 1, Color.LimeGreen)},
              {"All Incoming (Ref)",new CombatantData.DamageTypeDef("All Incoming (Ref)", 0, Color.Black)}
            };
            CombatantData.SwingTypeToDamageTypeDataLinksOutgoing = new SortedDictionary<int, List<string>>
            {
              {1, new List<string> { OUT_DAMAGE, "Outgoing Damage" } },
              {3, new List<string> { OUT_HEAL, "Outgoing Heal (Out)" } },
            };
            CombatantData.SwingTypeToDamageTypeDataLinksIncoming = new SortedDictionary<int, List<string>>
            {
              {1, new List<string> { INC_DAMAGE } },
              {2, new List<string> { INC_DAMAGE } },
              {3, new List<string> { "Healed (Inc)" } },
              {4, new List<string> { "Healed (Inc)" } },
            };

            CombatantData.DamageSwingTypes = new List<int> { 1, 2 };
            CombatantData.HealingSwingTypes = new List<int> { 3, 4 };

            CombatantData.DamageTypeDataNonSkillDamage = OUT_DAMAGE;
            CombatantData.DamageTypeDataOutgoingDamage = "Outgoing Damage";
            CombatantData.DamageTypeDataOutgoingHealing = "Outgoing Heal (Out)";
            CombatantData.DamageTypeDataIncomingDamage = INC_DAMAGE;
            CombatantData.DamageTypeDataIncomingHealing = "Healed (Inc)";

            try
            {
                // Columns "Zone View Options"
                EncounterData.ColumnDefs.Add("Healed", new EncounterData.ColumnDef("Healed", false, "INT", "Healed", (Data) => { return GetHealData(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetHealData(Data, SecretLanguage.Aegis).ToString(); }));
                EncounterData.ColumnDefs.Add("AegisDmg", new EncounterData.ColumnDef("AegisDmg", false, "INT", "AegisDmg", (Data) => { return GetSpecialHitData(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitData(Data, SecretLanguage.Aegis).ToString(); }));
                EncounterData.ColumnDefs.Add("AegisDPS", new EncounterData.ColumnDef("AegisDPS", false, "FLOAT", "AegisDPS", (Data) => { return CalcDataPS(GetSpecialHitData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(GetFloatCommas()); }, (Data) => { return CalcDataPS(GetSpecialHitData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(); }));
                EncounterData.ColumnDefs.Add("AegisHeal", new EncounterData.ColumnDef("AegisHeal", false, "INT", "AegisHeal", (Data) => { return GetSpecialHealData(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHealData(Data, SecretLanguage.Aegis).ToString(); }));
                EncounterData.ColumnDefs.Add("AegisHPS", new EncounterData.ColumnDef("AegisHPS", false, "FLOAT", "AegisHPS", (Data) => { return CalcDataPS(GetSpecialHealData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(GetFloatCommas()); }, (Data) => { return CalcDataPS(GetSpecialHealData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(); }));

                // Columns "Encounter View Options"
                CombatantData.ColumnDefs.Remove("Cures");
                CombatantData.ColumnDefs.Remove("PowerDrain");
                CombatantData.ColumnDefs.Remove("PowerReplenish");
                CombatantData.ColumnDefs.Remove("Threat +/-");
                CombatantData.ColumnDefs.Remove("ThreatDelta");

                CombatantData.ExportVariables.Add("Glance%", new CombatantData.TextExportFormatter("Glance%", "Glance%", "Attacks that glance in %", (Data, Extra) => { return CombatantFormatSwitch(Data, "Glance%", usCulture); }));
                CombatantData.ExportVariables.Add("Penetration%", new CombatantData.TextExportFormatter("Penetration%", "Penetration%", "Attacks that penetrate in %", (Data, Extra) => { return CombatantFormatSwitch(Data, "Penetration%", usCulture); }));
                CombatantData.ExportVariables.Add("Blocked%", new CombatantData.TextExportFormatter("Blocked%", "Blocked%", "Attacks that are blocked in %", (Data, Extra) => { return CombatantFormatSwitch(Data, "Blocked%", usCulture); }));

                CombatantData.ColumnDefs.Add("AegisDmg", new CombatantData.ColumnDef("AegisDmg", false, "INT", "AegisDmg", (Data) => { return GetSpecialHitData(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitData(Data, SecretLanguage.Aegis).ToString(); }, (Left, Right) => { return GetSpecialHitData(Left, SecretLanguage.Aegis).CompareTo(GetSpecialHitData(Right, SecretLanguage.Aegis)); }));
                CombatantData.ColumnDefs.Add("AegisDPS", new CombatantData.ColumnDef("AegisDPS", false, "FLOAT", "AegisDPS", (Data) => { return CalcDataPS(GetSpecialHitData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(GetFloatCommas()); }, (Data) => { return CalcDataPS(GetSpecialHitData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(); }, (Left, Right) => { return CalcDataPS(GetSpecialHitData(Left, SecretLanguage.Aegis), Left.Duration.TotalSeconds).CompareTo(CalcDataPS(GetSpecialHitData(Right, SecretLanguage.Aegis), Right.Duration.TotalSeconds)); }));
                CombatantData.ColumnDefs.Add("AegisHeal", new CombatantData.ColumnDef("AegisHeal", false, "INT", "AegisHeal", (Data) => { return GetSpecialHealData(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHealData(Data, SecretLanguage.Aegis).ToString(); }, (Left, Right) => { return GetSpecialHealData(Left, SecretLanguage.Aegis).CompareTo(GetSpecialHealData(Right, SecretLanguage.Aegis)); }));
                CombatantData.ColumnDefs.Add("AegisHPS", new CombatantData.ColumnDef("AegisHPS", false, "FLOAT", "AegisHPS", (Data) => { return CalcDataPS(GetSpecialHealData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(GetFloatCommas()); }, (Data) => { return CalcDataPS(GetSpecialHealData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(); }, (Left, Right) => { return CalcDataPS(GetSpecialHealData(Left, SecretLanguage.Aegis), Left.Duration.TotalSeconds).CompareTo(CalcDataPS(GetSpecialHealData(Right, SecretLanguage.Aegis), Right.Duration.TotalSeconds)); }));

                // Columns "Combatant View Options"
                DamageTypeData.ColumnDefs.Add("GlanceHits", new DamageTypeData.ColumnDef("GlanceHits", false, "INT3", "GlanceHits", (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Glancing).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Glancing).ToString(); }));
                DamageTypeData.ColumnDefs.Add("Glance%", new DamageTypeData.ColumnDef("Glance%", false, "VARCHAR(8)", "GlancePerc", (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Glancing).ToString("0'%"); }, (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Glancing).ToString(); }));
                DamageTypeData.ColumnDefs.Add("PenetrationHits", new DamageTypeData.ColumnDef("PenetrationHits", false, "INT3", "PenetrationHits", (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Penetrated).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Penetrated).ToString(); }));
                DamageTypeData.ColumnDefs.Add("Penetration%", new DamageTypeData.ColumnDef("Penetration%", false, "VARCHAR(8)", "PenetrationPerc", (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Penetrated).ToString("0'%"); }, (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Penetrated).ToString(); }));
                DamageTypeData.ColumnDefs.Add("BlockedHits", new DamageTypeData.ColumnDef("BlockedHits", false, "INT3", "BlockedHits", (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Blocked).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Blocked).ToString(); }));
                DamageTypeData.ColumnDefs.Add("Blocked%", new DamageTypeData.ColumnDef("Blocked%", false, "VARCHAR(8)", "BlockedPerc", (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Blocked).ToString("0'%"); }, (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Blocked).ToString(); }));
                DamageTypeData.ColumnDefs.Add("AegisHits", new DamageTypeData.ColumnDef("AegisHits", false, "INT3", "AegisHits", (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Aegis).ToString(); }));
                DamageTypeData.ColumnDefs.Add("AegisDmg", new DamageTypeData.ColumnDef("AegisDmg", false, "INT", "AegisDmg", (Data) => { return GetSpecialHitData(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitData(Data, SecretLanguage.Aegis).ToString(); }));
                DamageTypeData.ColumnDefs.Add("AegisDPS", new DamageTypeData.ColumnDef("AegisDPS", false, "FLOAT", "AegisDPS", (Data) => { return CalcDataPS(GetSpecialHitData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(GetFloatCommas()); }, (Data) => { return CalcDataPS(GetSpecialHitData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(); }));
                DamageTypeData.ColumnDefs.Add("AegisMismatch", new DamageTypeData.ColumnDef("AegisMismatch", false, "INT3", "AegisMismatch", (Data) => { return GetSpecialHitMiss(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitMiss(Data, SecretLanguage.Aegis).ToString(); }));
                DamageTypeData.ColumnDefs.Add("AegisMismatch%", new DamageTypeData.ColumnDef("AegisMismatch%", false, "VARCHAR(8)", "AegisMismatchPerc", (Data) => { return GetSpecialHitMissPerc(Data, SecretLanguage.Aegis).ToString("0'%"); }, (Data) => { return GetSpecialHitMissPerc(Data, SecretLanguage.Aegis).ToString(); }));


                // Columns "DamageType View Options"
                AttackType.ColumnDefs.Add("GlanceHits", new AttackType.ColumnDef("GlanceHits", false, "INT3", "GlanceHits", (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Glancing).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Glancing).ToString(); }, (Left, Right) => { return GetSpecialHitCount(Left, SecretLanguage.Glancing).CompareTo(GetSpecialHitCount(Right, SecretLanguage.Glancing)); }));
                AttackType.ColumnDefs.Add("Glance%", new AttackType.ColumnDef("Glance%", false, "VARCHAR(8)", "GlancePerc", (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Glancing).ToString("0'%"); }, (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Glancing).ToString(); }, (Left, Right) => { return GetSpecialHitPerc(Left, SecretLanguage.Glancing).CompareTo(GetSpecialHitPerc(Right, SecretLanguage.Glancing)); }));
                AttackType.ColumnDefs.Add("PenetrationHits", new AttackType.ColumnDef("PenetrationHits", false, "INT3", "PenetrationHits", (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Penetrated).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Penetrated).ToString(); }, (Left, Right) => { return GetSpecialHitCount(Left, SecretLanguage.Penetrated).CompareTo(GetSpecialHitCount(Right, SecretLanguage.Penetrated)); }));
                AttackType.ColumnDefs.Add("Penetration%", new AttackType.ColumnDef("Penetration%", false, "VARCHAR(8)", "PenetrationPerc", (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Penetrated).ToString("0'%"); }, (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Penetrated).ToString(); }, (Left, Right) => { return GetSpecialHitPerc(Left, SecretLanguage.Penetrated).CompareTo(GetSpecialHitPerc(Right, SecretLanguage.Penetrated)); }));
                AttackType.ColumnDefs.Add("BlockedHits", new AttackType.ColumnDef("BlockedHits", false, "INT3", "BlockedHits", (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Blocked).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Blocked).ToString(); }, (Left, Right) => { return GetSpecialHitCount(Left, SecretLanguage.Blocked).CompareTo(GetSpecialHitCount(Right, SecretLanguage.Blocked)); }));
                AttackType.ColumnDefs.Add("Blocked%", new AttackType.ColumnDef("Blocked%", false, "VARCHAR(8)", "BlockedPerc", (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Blocked).ToString("0'%"); }, (Data) => { return GetSpecialHitPerc(Data, SecretLanguage.Blocked).ToString(); }, (Left, Right) => { return GetSpecialHitPerc(Left, SecretLanguage.Blocked).CompareTo(GetSpecialHitPerc(Right, SecretLanguage.Blocked)); }));
                AttackType.ColumnDefs.Add("AegisHits", new AttackType.ColumnDef("AegisHits", false, "INT3", "AegisHits", (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitCount(Data, SecretLanguage.Aegis).ToString(); }, (Left, Right) => { return GetSpecialHitCount(Left, SecretLanguage.Aegis).CompareTo(GetSpecialHitCount(Right, SecretLanguage.Aegis)); }));
                AttackType.ColumnDefs.Add("AegisDmg", new AttackType.ColumnDef("AegisDmg", false, "INT", "AegisDmg", (Data) => { return GetSpecialHitData(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitData(Data, SecretLanguage.Aegis).ToString(); }, (Left, Right) => { return GetSpecialHitData(Left, SecretLanguage.Aegis).CompareTo(GetSpecialHitData(Right, SecretLanguage.Aegis)); }));
                AttackType.ColumnDefs.Add("AegisDPS", new AttackType.ColumnDef("AegisDPS", false, "FLOAT", "AegisDPS", (Data) => { return CalcDataPS(GetSpecialHitData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(GetFloatCommas()); }, (Data) => { return CalcDataPS(GetSpecialHitData(Data, SecretLanguage.Aegis), Data.Duration.TotalSeconds).ToString(); }, (Left, Right) => { return CalcDataPS(GetSpecialHitData(Left, SecretLanguage.Aegis), Left.Duration.TotalSeconds).CompareTo(CalcDataPS(GetSpecialHitData(Right, SecretLanguage.Aegis), Right.Duration.TotalSeconds)); }));
                AttackType.ColumnDefs.Add("AegisMismatch", new AttackType.ColumnDef("AegisMismatch", false, "INT3", "AegisMismatch", (Data) => { return GetSpecialHitMiss(Data, SecretLanguage.Aegis).ToString(GetIntCommas()); }, (Data) => { return GetSpecialHitMiss(Data, SecretLanguage.Aegis).ToString(); }, (Left, Right) => { return GetSpecialHitMiss(Left, SecretLanguage.Aegis).CompareTo(GetSpecialHitMiss(Right, SecretLanguage.Aegis)); }));
                AttackType.ColumnDefs.Add("AegisMismatch%", new AttackType.ColumnDef("AegisMismatch%", false, "VARCHAR(8)", "AegisMismatchPerc", (Data) => { return GetSpecialHitMissPerc(Data, SecretLanguage.Aegis).ToString("0'%"); }, (Data) => { return GetSpecialHitMissPerc(Data, SecretLanguage.Aegis).ToString(); }, (Left, Right) => { return GetSpecialHitMissPerc(Left, SecretLanguage.Aegis).CompareTo(GetSpecialHitMissPerc(Right, SecretLanguage.Aegis)); }));
            }
            catch (ArgumentException)
            { }

            ActGlobals.oFormActMain.ValidateLists();
            ActGlobals.oFormActMain.ValidateTableSetup();
        }

        private bool IsEncounterSelected()
        {
            return (ActGlobals.oFormActMain.MainTreeView.SelectedNode != null &&
                ActGlobals.oFormActMain.MainTreeView.SelectedNode.Parent != null &&
                ActGlobals.oFormActMain.MainTreeView.SelectedNode.Parent.Parent == null);
        }

        private void ExportEncounterToScript(object sender, EventArgs e)
        {
            if (IsEncounterSelected())
            {
                var encounter = ActGlobals.oFormActMain.ZoneList[ActGlobals.oFormActMain.MainTreeView.SelectedNode.Parent.Index].Items[ActGlobals.oFormActMain.MainTreeView.SelectedNode.Index];
                CultureInfo usCulture = new CultureInfo("en-US");
                Dictionary<string, Dictionary<string, string>> lines = new Dictionary<string, Dictionary<string, string>>();
                Dictionary<string, Dictionary<string, Dictionary<string, string>>> embedLines = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();
                SortedDictionary<string, string> displayOrder = new SortedDictionary<string, string>();
                SortedDictionary<string, string> displayOrderHealers = new SortedDictionary<string, string>();
                SortedDictionary<string, string> displayOrderTanks = new SortedDictionary<string, string>();
                SortedDictionary<string, SortedDictionary<string, string>> embedDisplayOrder = new SortedDictionary<string, SortedDictionary<string, string>>();
                bool isImport = "Import Zone".Equals(encounter.ZoneName);
                BuildScriptOutput(encounter, usCulture, lines, embedLines, displayOrder, embedDisplayOrder, displayOrderHealers, displayOrderTanks);
                GenerateHtmlScript(encounter, usCulture, lines, embedLines, displayOrder, embedDisplayOrder);
                GenerateChatScript(encounter, usCulture, lines, displayOrder, displayOrderHealers, displayOrderTanks, isImport);
            }
        }

        private string GetIntCommas()
        {
            return ActGlobals.mainTableShowCommas ? "#,0" : "0";
        }

        protected string GetFloatCommas()
        {
            return ActGlobals.mainTableShowCommas ? "#,0.00" : "0.00";
        }

        private string GetScriptKey(long ToConvert)
        {
            return (9999999999 - ToConvert).ToString("0000000000");
        }

        private string GetScriptKey(long dmg, long heal, string name)
        {
            return GetScriptKey(dmg) + GetScriptKey(heal) + name;
        }

        private string GetScriptKey(string dmg, string heal, string name)
        {
            return dmg + heal + name;
        }

        private string GetScriptFolder()
        {
            try
            {
                return Path.Combine(Path.GetDirectoryName(ActGlobals.oFormActMain.LogFilePath), "Scripts");
            }
            catch (Exception e)
            {
                throw new Exception("Failed to get script folder from directory: " + ActGlobals.oFormActMain.LogFilePath, e);
            }
        }

        private void AppendHtmlLine(StringBuilder line, Dictionary<string, string> entry, string rowClass, string nameStr, bool hidden, int embed_row, int IsMultipleDeath)
        {
            string TrOnClick = "even".Equals(rowClass) || "odd".Equals(rowClass) ? rowClass + "\" onclick=\"toggle('embed" + embed_row + "');"
                                                                                   + ("0".Equals(entry["dmg"]) && "0".Equals(entry["healing"]) ? "" : "this.style.background=(this.style.background==''?'#FCFFB0':'');")
                                                                                 : rowClass;
            string hiddenStr = hidden ? " style=\"display: none;\"" : "";
            string death = ("0".Equals(entry["death"]) || IsMultipleDeath == 1) ? "" : " - " + (SecretLanguage.Killing).ToLower();
            string duration = "".Equals(entry["duration"]) ? "" : " - ( " + entry["duration"] + death + " )";

            string HTMLine = "<tr class=\"" + TrOnClick + "\"" + hiddenStr + "><td class=\"col1\">"
                           + ("".Equals(entry["duration"]) ? "<b>" + nameStr + "</b>" : "<strong>" + nameStr + "</strong>")
                           + duration + "</td>";
            HTMLine += "0".Equals(entry["dmg"]) ? "<td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>"
                                              : "<td>" + entry["dmg"] + " - " + entry["dps%"] + "</td><td>"
                                                + ("".Equals(duration) ? "" + "</td><td>" + ("0%".Equals(entry["pen%"]) ? "" : entry["pen%"])
                                                                          + "</td><td>" + ("0%".Equals(entry["crit%"]) ? "" : entry["crit%"])
                                                                          + "</td><td>" + ("0%".Equals(entry["glance%"]) ? "" : entry["glance%"])
                                                                          + "</td><td>" + ("0%".Equals(entry["block%"]) ? "" : entry["block%"])
                                                                          + "</td><td>" + ("0%".Equals(entry["evade%"]) ? "" : entry["evade%"])
                                                                          + "</td><td>" + ("0".Equals(entry["aegisdmg"]) ? "" : entry["aegisdmg"])
                                                                          + "</td><td>" + ("0.00".Equals(entry["aegisdps"]) ? "" : entry["aegisdps"])
                                                                          + "</td><td>" + ("0%".Equals(entry["aegismismatch%"]) ? "" : entry["aegismismatch%"])
                                                                          + "</td>"
                                                                        : entry["dps"] + "</td><td>" + entry["pen%"]
                                                                          + "</td><td>" + entry["crit%"]
                                                                          + "</td><td>" + entry["glance%"]
                                                                          + "</td><td>" + entry["block%"]
                                                                          + "</td><td>" + entry["evade%"]
                                                                          + "</td><td>" + ("0".Equals(entry["aegisdmg"]) ? "" : entry["aegisdmg"])
                                                                          + "</td><td>" + ("0.00".Equals(entry["aegisdps"]) ? "" : entry["aegisdps"])
                                                                          + "</td><td>" + ("0%".Equals(entry["aegismismatch%"]) ? "" : entry["aegismismatch%"])
                                                                          + "</td>");
            HTMLine += "0".Equals(entry["healing"]) ? "<td class=\"heal\"></td><td></td><td></td><td></td><td></td>"
                                                  : "<td class=\"heal\">" + entry["healing"] + " - " + entry["hps%"]
                                                    + "</td><td>" + entry["hps"]
                                                    + "</td><td>" + entry["healcrit%"] + "</td>"
                                                    + "</td><td>" + ("0".Equals(entry["aegisheal"]) ? "" : entry["aegisheal"])
                                                    + "</td><td>" + ("0.00".Equals(entry["aegishps"]) ? "" : entry["aegishps"]);
            HTMLine += "</tr>";
            line.AppendFormat(HTMLine);
        }

        void oFormActMain_OnCombatStart(bool isImport, CombatToggleEventArgs encounterInfo)
        {
            lock (locker)
            {
                lastCombatStart = ActGlobals.oFormActMain.LastKnownTime;

                if (buffData == null)
                {
                    buffData = new Dictionary<string, BuffData>();
                }
                else
                {
                    List<string> buffDeleteList = new List<string>();
                    foreach (var buff in buffData)
                    {
                        if (buff.Value.IsRunning())
                        {
                            if (buff.Value.StartTime.CompareTo(lastCombatEnd) < 1)
                            {
                                buff.Value.Restart(ActGlobals.oFormActMain.LastKnownTime);
                            }
                        }
                        else
                        {
                            buffDeleteList.Add(buff.Key);
                        }
                    }

                    foreach (var buff in buffDeleteList)
                    {
                        buffData.Remove(buff);
                    }
                }
            }
        }

        void oFormActMain_OnCombatEnd(bool isImport, CombatToggleEventArgs encounterInfo)
        {
            lock (locker)
            {
                lastCombatEnd = encounterInfo.encounter.EndTime;

                if (allies != null && allies.Count > 0)
                {
                    List<CombatantData> localAllies = new List<CombatantData>(allies.Count);
                    foreach (var name in allies)
                    {
                        var combatant = encounterInfo.encounter.GetCombatant(name);
                        if (combatant != null)
                        {
                            localAllies.Add(encounterInfo.encounter.GetCombatant(name));
                        }
                    }

                    encounterInfo.encounter.SetAllies(localAllies);
                }

                string charName = ConvertCharName(SecretLanguage.You);
                if (encounterInfo.encounter.GetCombatant(charName) != null)
                {
                    string bossName = encounterInfo.encounter.GetStrongestEnemy(null);
                    if (bossName.IndexOf("_self_") > 0) return; // Prevent _self_-encounters due to dot after fights
                    encounterInfo.encounter.Title = bossName;
                }
                else
                {
                    return;
                }

                Dictionary<string, BuffData> buffDataCopy = new Dictionary<string, BuffData>();
                foreach (var buff in buffData)
                {
                    if (buff.Value.Changed)
                    {
                        BuffData copy = buff.Value.Copy();
                        copy.Stop(encounterInfo.encounter.EndTime);
                        if (copy.TotalTime > encounterInfo.encounter.Duration.TotalMilliseconds)
                        {
                            copy = new BuffData();
                            copy.Start(encounterInfo.encounter.StartTime);
                            copy.Stop(encounterInfo.encounter.EndTime);
                        }

                        buffDataCopy.Add(buff.Key, copy);
                    }
                }

                encounterInfo.encounter.Tags[BUFFS] = buffDataCopy;
            }

            if (checkBox_DontExportShortEnc.Checked && encounterInfo.encounter.Duration.TotalSeconds < 5)
            {
                return;
            }

            CultureInfo usCulture = new CultureInfo("en-US");
            Dictionary<string, Dictionary<string, string>> lines = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> embedLines = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();
            SortedDictionary<string, string> displayOrder = new SortedDictionary<string, string>();
            SortedDictionary<string, string> displayOrderHealers = new SortedDictionary<string, string>();
            SortedDictionary<string, string> displayOrderTanks = new SortedDictionary<string, string>();
            SortedDictionary<string, SortedDictionary<string, string>> embedDisplayOrder = new SortedDictionary<string, SortedDictionary<string, string>>();
            if (checkBox_ExportHtml.Checked || checkBox_ExportScript.Checked)
            {
                BuildScriptOutput(encounterInfo.encounter, usCulture, lines, embedLines, displayOrder, embedDisplayOrder, displayOrderHealers, displayOrderTanks);

                if (checkBox_ExportHtml.Checked)
                {
                    GenerateHtmlScript(encounterInfo.encounter, usCulture, lines, embedLines, displayOrder, embedDisplayOrder);
                }

                if (checkBox_ExportScript.Checked)
                {
                    GenerateChatScript(encounterInfo.encounter, usCulture, lines, displayOrder, displayOrderHealers, displayOrderTanks, isImport);
                }
            }
        }

        private string AddString(string Complete, string Add)
        {
            string Result = "";
            if (Complete != "") Result = Complete + ", ";
            Result += Add;
            return Result;
        }

        private void GenerateChatScript(EncounterData encounter, CultureInfo usCulture, Dictionary<string, Dictionary<string, string>> lines, SortedDictionary<string, string> displayOrder, SortedDictionary<string, string> displayOrderHealers, SortedDictionary<string, string> displayOrderTanks, bool isImport = false)
        {
            string scriptFolder = GetScriptFolder();
            if (Directory.Exists(scriptFolder))
            {
                const int NameLength = 7;
                StringBuilder line = new StringBuilder();
                StringBuilder lineSplit = new StringBuilder();
                bool exportColored = checkBox_ExportColored.Checked;
                string title = encounter.Title.Replace("\"", "&quot;");
                string combat_duration = encounter.Duration.TotalSeconds > 599 ? encounter.DurationS : encounter.DurationS.Substring(1, 4);
                string hitpoints = encounter.Damage.ToString("#,##0", usCulture);
                string hdrOutput = "";
                string hdrDamage = exportColored ? "<font face=LARGE_BOLD color=red>--- Damage</font><br>" : "--- Damage<br>";
                string hdrHeal = exportColored ? "<font face=LARGE_BOLD color=#25d425>--- Heal</font><br>" : "--- Heal<br>";
                string hdrTank = exportColored ? "<font face=LARGE_BOLD color=#09e4ea>--- Tank</font><br>" : "--- Tank<br>";
                string hdrMax = exportColored ? "<font face=LARGE_BOLD color=#be09cc>--- Max</font><br>" : "--- Max<br>";
                string lineReduced = "({0} more)<br>";
                List<String> lineDamage = new List<String>();
                List<String> lineHeal = new List<String>();
                List<String> lineTank = new List<String>();
                List<String> lineMax = new List<String>();

                // Gather data for damage, heal, tank and max
                long aegis_hp = GetSpecialHitData(encounter, SecretLanguage.Aegis);
                if (aegis_hp != 0)
                {
                    hitpoints += " [" + aegis_hp.ToString("#,##0", usCulture) + "]";
                }
                string heading = string.Format("total: {2} dmg, {0} dps in {1}", encounter.DPS.ToString("#,##0", usCulture), combat_duration, hitpoints);
                hdrOutput = string.Format("<a href=\"text://<div align=center><font face=HEADLINE color=red>{0}</font><br><font face=HUGE color=#FF6600>{1}</font></div><br><font face=LARGE>", title, heading);

                if (IsExportFieldSet("dps") || IsExportFieldSet("crit%") || IsExportFieldSet("pen%") || IsExportFieldSet("glance%") || IsExportFieldSet("block%") || IsExportFieldSet("evade%") || (IsExportFieldSet("aegismismatch%") && (aegis_hp != 0)))
                {
                    foreach (var name in displayOrder.Values)
                    {
                        Dictionary<string, string> entry = lines[name];
                        if ("0".Equals(entry["dmg"]))
                        {
                            continue;
                        }

                        string tempName = name;
                        if (checkBox_LimitNames.Checked)
                        {
                            if (name.Length > NameLength)
                            {
                                tempName = name.Substring(0, NameLength);
                            }
                        }

                        string combatant = string.Format("{0}", tempName);
                        combatant += AddChatScriptField(entry, "dps");
                        combatant += AddChatScriptField(entry, "pen%");
                        combatant += AddChatScriptField(entry, "crit%");
                        combatant += AddChatScriptField(entry, "glance%");
                        combatant += AddChatScriptField(entry, "block%");
                        combatant += AddChatScriptField(entry, "evade%");
                        if (aegis_hp != 0) combatant += AddChatScriptField(entry, "aegismismatch%");
                        lineDamage.Add(combatant + "<br>");
                    }
                }

                if (IsExportFieldSet("hps") || IsExportFieldSet("healcrit%"))
                {
                    foreach (var name in displayOrderHealers.Values)
                    {
                        Dictionary<string, string> entry = lines[name];
                        if ("0".Equals(entry["healing"]))
                        {
                            continue;
                        }

                        string tempName = name;
                        if (checkBox_LimitNames.Checked)
                        {
                            if (name.Length > NameLength)
                            {
                                tempName = name.Substring(0, NameLength);
                            }
                        }

                        string combatant = string.Format("{0}", tempName);
                        combatant += AddChatScriptField(entry, "hps");
                        combatant += AddChatScriptField(entry, "healcrit%");
                        lineHeal.Add(combatant + "<br>");
                    }
                }

                if (IsExportFieldSet("TakenDamage") || IsExportFieldSet("TakenCrit%") || IsExportFieldSet("TakenPen%") || IsExportFieldSet("TakenGlance%") || IsExportFieldSet("TakenBlock%") || IsExportFieldSet("TakenEvade%"))
                {
                    foreach (var name in displayOrderTanks.Values)
                    {
                        Dictionary<string, string> entry = lines[name];
                        if ("0".Equals(entry["takendamage"]))
                        {
                            continue;
                        }

                        string tempName = name;
                        if (checkBox_LimitNames.Checked)
                        {
                            if (name.Length > NameLength)
                            {
                                tempName = name.Substring(0, NameLength);
                            }
                        }

                        if (!isImport && entry.ContainsKey("death") && !"0".Equals(entry["death"]))
                        {
                            tempName += " †";
                        }

                        string combatant = string.Format("{0}", tempName);
                        combatant += AddChatScriptField(entry, "takendamage");
                        combatant += AddChatScriptField(entry, "takenpen%");
                        combatant += AddChatScriptField(entry, "takencrit%");
                        combatant += AddChatScriptField(entry, "takenglance%");
                        combatant += AddChatScriptField(entry, "takenblock%");
                        combatant += AddChatScriptField(entry, "takenevade%");
                        lineTank.Add(combatant + "<br>");
                    }
                }

                if (encounter.Damage > 0)
                {
                    string maxHit;
                    string maxHeal;
                    string maxDamage;

                    GetMaxText(encounter.GetAllies(), true, usCulture, out maxHit, out maxHeal, out maxDamage);
                    if (maxHit.Length > 0)
                    {
                        lineMax.Add(string.Format("Hit: {0}<br>", maxHit));
                    }

                    if (maxHeal.Length > 0)
                    {
                        lineMax.Add(string.Format("Heal: {0}<br>", maxHeal));
                    }

                    if (maxDamage.Length > 0)
                    {
                        lineMax.Add(string.Format("Incoming: {0}<br>", maxDamage));
                    }
                }

                string Expl = "";
                if (checkBox_ExportShowLegend.Checked)
                {
                    if (aegis_hp != 0) Expl = AddString(Expl, "[aegis]");
                    if (IsExportFieldSet("pen%") || IsExportFieldSet("takenpen%")) Expl = AddString(Expl, "p=pen");
                    if (IsExportFieldSet("crit%") || IsExportFieldSet("takencrit%") || IsExportFieldSet("healcrit%")) Expl = AddString(Expl, "c=crit");
                    if (IsExportFieldSet("glance%") || IsExportFieldSet("takenglance%")) Expl = AddString(Expl, "g=glance");
                    if (IsExportFieldSet("block%") || IsExportFieldSet("takenblock%")) Expl = AddString(Expl, "b=block");
                    if (IsExportFieldSet("evade%") || IsExportFieldSet("takenevade%")) Expl = AddString(Expl, "e=evade");
                    if (IsExportFieldSet("aegismismatch%") && (aegis_hp != 0)) Expl = AddString(Expl, "a=aegis mismatch");
                    Expl = "<br><font face=HUGE color=#828282>" + Expl + "</font>";
                }

                // Build page
                line.Append(hdrOutput);
                line.Append("<div>");
                if (lineDamage.Count > 0)
                {
                    line.Append(hdrDamage);
                    line.Append("{0}</div><br><div>");
                }
                if (lineHeal.Count > 0)
                {
                    line.Append(hdrHeal);
                    line.Append("{1}</div><br><div>");
                }
                if (lineTank.Count > 0)
                {
                    line.Append(hdrTank);
                    line.Append("{2}</div><br><div>");
                }
                AppendLineMax(lineMax, hdrMax, line);
                line.Append("</div>");
                line.Append("</font>");
                line.AppendFormat(Expl + "\">{0} - {1}</a>", title, heading);

                int DmgLen = GetTotalLength(lineDamage);
                int HealLen = GetTotalLength(lineHeal);
                int TankLen = GetTotalLength(lineTank);
                int MaxLen = GetTotalLength(lineMax);

                int limitDamage = DmgLen;
                int limitHeal = HealLen;
                int limitTank = TankLen;

                bool scriptToLong = line.Length + DmgLen + HealLen + TankLen > CHAT_LIMIT;

                if (scriptToLong)
                {
                    int countData = CHAT_LIMIT - line.Length;
                    float Ratio = (float)countData / (float)(DmgLen + HealLen + TankLen);

                    limitDamage = (int)Math.Round(DmgLen * Ratio);
                    limitHeal = (int)Math.Round(HealLen * Ratio);
                    limitTank = (int)Math.Round(TankLen * Ratio);
                }

                string outDamage = "";
                string outHeal = "";
                string outTank = "";

                for (var i = 0; i < lineDamage.Count; i++)
                {
                    if (outDamage.Length + lineDamage[i].Length <= limitDamage)
                    {
                        outDamage += lineDamage[i];
                    }
                    else
                    {
                        outDamage += string.Format(lineReduced, lineDamage.Count - i);
                        break;
                    }
                }

                for (var i = 0; i < lineHeal.Count; i++)
                {
                    if (outHeal.Length + lineHeal[i].Length <= limitHeal)
                    {
                        outHeal += lineHeal[i];
                    }
                    else
                    {
                        outHeal += string.Format(lineReduced, lineHeal.Count - i);
                        break;
                    }
                }

                for (var i = 0; i < lineTank.Count; i++)
                {
                    if (outTank.Length + lineTank[i].Length <= limitTank)
                    {
                        outTank += lineTank[i];
                    }
                    else
                    {
                        outTank += string.Format(lineReduced, lineTank.Count - i);
                        break;
                    }
                }

                string output = string.Format(line.ToString(), outDamage, outHeal, outTank);

                if (scriptToLong && checkBox_ExportSplit.Checked)
                {
                    // actchatsplit chat script
                    lineSplit.Append("<font color=red>[ ").Append(title).Append(" - ").Append(heading).Append(" ]</font>").AppendLine();

                    string linkStart = "<a href=\"text://<div align=center><font face=HEADLINE color=red>";
                    string linkCenter = "</font><br><font face=HUGE color=#FF6600>";
                    string linkEnd = "</font></div><br><font face=LARGE>";
                    string tagDivStart = "<div>";
                    string tagDivEnd = "</div><br>";
                    string tagFontEnd = "</font>";

                    string linkDamage = "\">Damage Report</a>";
                    string linkHeal = "\">Heal Report</a>";
                    string linkTank = "\">Tank Report</a>";

                    int scriptBaseLength = linkStart.Length + linkCenter.Length + linkEnd.Length + tagDivStart.Length + tagDivEnd.Length + tagFontEnd.Length + title.Length + heading.Length + Expl.Length + MaxLen + hdrMax.Length;
                    int scriptDamageLength = scriptBaseLength + hdrDamage.Length + linkDamage.Length;
                    int scriptHealLength = scriptBaseLength + hdrHeal.Length + linkHeal.Length;
                    int scriptTankLength = scriptBaseLength + hdrTank.Length + linkTank.Length;

                    // Damage Report
                    if (lineDamage.Count > 0)
                    {
                        lineSplit.Append(linkStart).Append(title).Append(linkCenter).Append(heading).Append(linkEnd);
                        lineSplit.Append(tagDivStart).Append(hdrDamage);
                        if (scriptDamageLength + DmgLen > CHAT_LIMIT)
                        {
                            AppendLinesReduced(lineDamage, lineSplit, CHAT_LIMIT - scriptDamageLength, lineReduced);
                        }
                        else
                        {
                            AppendLines(lineDamage, lineSplit);
                        }
                        lineSplit.Append(tagDivEnd);
                        AppendLineMax(lineMax, hdrMax, lineSplit);
                        lineSplit.Append(tagFontEnd).Append(Expl).Append(linkDamage).AppendLine();
                    }

                    // Heal Report
                    if (lineHeal.Count > 0)
                    {
                        lineSplit.Append(linkStart).Append(title).Append(linkCenter).Append(heading).Append(linkEnd);
                        lineSplit.Append(tagDivStart).Append(hdrHeal);
                        if (scriptHealLength + HealLen > CHAT_LIMIT)
                        {
                            AppendLinesReduced(lineHeal, lineSplit, CHAT_LIMIT - scriptHealLength, lineReduced);
                        }
                        else
                        {
                            AppendLines(lineHeal, lineSplit);
                        }
                        lineSplit.Append(tagDivEnd);
                        AppendLineMax(lineMax, hdrMax, lineSplit);
                        lineSplit.Append(tagFontEnd).Append(Expl).Append(linkHeal).AppendLine();
                    }

                    // Tank Report
                    if (lineTank.Count > 0)
                    {
                        lineSplit.Append(linkStart).Append(title).Append(linkCenter).Append(heading).Append(linkEnd);
                        lineSplit.Append(tagDivStart).Append(hdrTank);
                        if (scriptTankLength + TankLen > CHAT_LIMIT)
                        {
                            AppendLinesReduced(lineTank, lineSplit, CHAT_LIMIT - scriptTankLength, lineReduced);
                        }
                        else
                        {
                            AppendLines(lineTank, lineSplit);
                        }
                        lineSplit.Append(tagDivEnd);
                        AppendLineMax(lineMax, hdrMax, lineSplit);
                        lineSplit.Append(tagFontEnd).Append(Expl).Append(linkTank);
                    }
                }

                try
                {
                    using (TextWriter writer = new StreamWriter(Path.Combine(scriptFolder, "actchat"), false, Encoding.GetEncoding(1252)))
                    {
                        if (scriptToLong && checkBox_ExportSplit.Checked)
                        {
                            writer.WriteLine(lineSplit.ToString());
                        }
                        else
                        {
                            writer.WriteLine(output);
                        }
                    }
                    using (TextWriter writer = new StreamWriter(Path.Combine(scriptFolder, "acttell"), false, Encoding.GetEncoding(1252)))
                    {
                        writer.WriteLine(SecretLanguage.WhisperCmd + " %1 " + output);
                    }
                }
                catch (Exception)
                {
                    // Ignore errors writing the file.  Assume this is a security violation or such
                }
            }
        }

        private static int GetTotalLength(List<String> lines)
        {
            int sum = 0;
            foreach (var line in lines)
            {
                sum += line.Length;
            }
            return sum;
        }

        private void AppendLineMax(List<String> lineMax, String hdrMax, StringBuilder line) {
            if (lineMax.Count > 0)
            {
                line.Append(hdrMax);
                for (var i = 0; i < lineMax.Count; i++)
                {
                    line.Append(lineMax[i]);
                }
            }
        }

        private void AppendLines(List<String> lines, StringBuilder line) {
            for (var i = 0; i < lines.Count; i++)
            {
                line.Append(lines[i]);
            }
        }

        private void AppendLinesReduced(List<String> lines, StringBuilder line, int charLimit, String lineReduced) {
            int usedChars = 0;
            int lineReducedLength = lineReduced.Length;
            int lineLength = 0;

            for (var i = 0; i < lines.Count; i++)
            {
                lineLength = lines[i].Length;
                if (usedChars + lineLength + lineReducedLength <= charLimit)
                {
                    line.Append(lines[i]);
                    usedChars += lineLength;
                }
                else
                {
                    line.Append(string.Format(lineReduced, lines.Count - i));
                    break;
                }
            }
        }

        private void GetMaxText(List<CombatantData> allys, bool showName, CultureInfo usCulture, out string maxHit, out string maxHeal, out String maxDamage)
        {
            maxHit = "";
            maxHeal = "";
            maxDamage = "";

            long iMaxHit = 0;
            long iMaxHeal = 0;
            long iMaxDamage = 0;
            foreach (var ally in allys)
            {
                if (ally.Items.ContainsKey(ALL_OUTGOING) && ally.Items[ALL_OUTGOING].Items.ContainsKey(ALL))
                {
                    var swingList = ally.Items[ALL_OUTGOING].Items[ALL].Items;
                    for (int indx = 0; indx < swingList.Count; ++indx)
                    {
                        if ((swingList[indx].SwingType == (int)SwingTypeEnum.Melee || swingList[indx].SwingType == (int)SwingTypeEnum.NonMelee) && swingList[indx].Damage.Number > iMaxHit)
                        {
                            iMaxHit = swingList[indx].Damage.Number;
                            maxHit = string.Format("{0} ({1})", iMaxHit.ToString("#,##0", usCulture), swingList[indx].AttackType);
                            if (showName)
                            {
                                maxHit = string.Format("{0} - {1}", swingList[indx].Attacker, maxHit);
                            }
                        }

                        if (swingList[indx].SwingType == (int)SwingTypeEnum.Healing && swingList[indx].Damage.Number > iMaxHeal)
                        {
                            iMaxHeal = swingList[indx].Damage.Number;
                            maxHeal = string.Format("{0} ({1})", iMaxHeal.ToString("#,##0", usCulture), swingList[indx].AttackType);
                            if (showName)
                            {
                                maxHeal = string.Format("{0} - {1}", swingList[indx].Attacker, maxHeal);
                            }
                        }
                    }
                }

                if (ally.Items.ContainsKey(INC_DAMAGE) && ally.Items[INC_DAMAGE].Items.ContainsKey(ALL))
                {
                    var swingList = ally.Items[INC_DAMAGE].Items[ALL].Items;
                    for (int indx = 0; indx < swingList.Count; ++indx)
                    {
                        if (swingList[indx].Damage.Number > iMaxDamage)
                        {
                            iMaxDamage = swingList[indx].Damage.Number;
                            maxDamage = string.Format("{0} ({1})", iMaxDamage.ToString("#,##0", usCulture), swingList[indx].AttackType);
                            if (showName)
                            {
                                maxDamage = string.Format("{0} - {1}", ally.Name, maxDamage);
                            }
                        }
                    }
                }
            }
        }

        private string AddChatScriptField(Dictionary<string, string> entry, string fieldName)
        {
            if (IsExportFieldSet(fieldName))
            {
                string BeforeField = " - ";
                string AfterField = "";
                bool exportColored = checkBox_ExportColored.Checked;
                switch (fieldName)
                {
                    case "dps":
                        BeforeField += entry["dmg_script"] + " ";
                        if (entry["aegisdmg_script"] != "") BeforeField += (exportColored ? "<font face=LARGE color=#1cbcea>[" : "[") + entry["aegisdmg_script"] + (exportColored ? "]</font> " : "] ");
                        BeforeField += "(";
                        if (entry["aegisdmg_script"] != "") AfterField += (exportColored ? " <font face=LARGE color=#1cbcea>[" : " [") + entry["aegisdps"] + (exportColored ? "]</font>" : "]");
                        AfterField += " dps in " + entry["duration"] + ")";
                        break;
                    case "hps":
                        BeforeField += entry["healing_script"] + " ";
                        if (entry["aegisheal_script"] != "") BeforeField += (exportColored ? "<font face=LARGE color=#1cbcea>[" : "[") + entry["aegisheal_script"] + (exportColored ? "]</font> " : "] ");
                        BeforeField += "(";
                        if (entry["aegisheal_script"] != "") AfterField += (exportColored ? " <font face=LARGE color=#1cbcea>[" : " [") + entry["aegishps"] + (exportColored ? "]</font>" : "]");
                        AfterField += " hps in " + entry["healduration"] + ")";
                        break;
                    case "pen%":
                        AfterField = "p";
                        break;
                    case "crit%":
                        AfterField = "c";
                        break;
                    case "glance%":
                        AfterField = "g";
                        break;
                    case "block%":
                        AfterField = "b";
                        break;
                    case "healcrit%":
                        AfterField = "c";
                        break;
                    case "evade%":
                        AfterField = "e";
                        break;
                    case "aegismismatch%":
                        AfterField = "a";
                        break;
                    case "takendamage":
                        if (entry["aegisinc_script"] != "") AfterField += (exportColored ? " <font face=LARGE color=#1cbcea>[" : " [") + entry["aegisinc_script"] + (exportColored ? "]</font>" : "]");
                        AfterField += " (";
                        AfterField += entry["takendps"];
                        if (entry["aegisinc_script"] != "") AfterField += (exportColored ? " <font face=LARGE color=#1cbcea>[" : " [") + entry["takenaegisdps"] + (exportColored ? "]</font>" : "]");
                        AfterField += " idps in " + entry["takenduration"] + ")";
                        break;
                    case "takencrit%":
                        AfterField = "c";
                        break;
                    case "takenpen%":
                        AfterField = "p";
                        break;
                    case "takenglance%":
                        AfterField = "g";
                        break;
                    case "takenblock%":
                        AfterField = "b";
                        break;
                    case "takenevade%":
                        AfterField = "e";
                        break;
                }
                return string.Format("{0}{1}{2}", BeforeField, entry[fieldName], AfterField);
            }
            return "";
        }

        private void GenerateHtmlScript(EncounterData encounter, CultureInfo usCulture, Dictionary<string, Dictionary<string, string>> lines, Dictionary<string, Dictionary<string, Dictionary<string, string>>> embedLines, SortedDictionary<string, string> displayOrder, SortedDictionary<string, SortedDictionary<string, string>> embedDisplayOrder)
        {
            string scriptFolder = GetScriptFolder();
            if (Directory.Exists(scriptFolder))
            {
                StringBuilder line = new StringBuilder(HTML_HEADER);
                for (int indx = 1; indx <= displayOrder.Count; ++indx)
                {
                    line.AppendLine("tr.embed" + indx.ToString() + " {background:#fff;}"
                                  + " tr.embed" + indx.ToString() + ":hover {background:#fef0f1;}");
                }

                string combat_duration = encounter.Duration.TotalSeconds > 599 ? encounter.DurationS : encounter.DurationS.Substring(1, 4);
                string hitpoints = encounter.Damage.ToString("#,##0", usCulture);
                long aegis_hp = GetSpecialHitData(encounter, SecretLanguage.Aegis);
                if (aegis_hp != 0)
                {
                    hitpoints += " [" + aegis_hp.ToString("#,##0", usCulture) + "]";
                }

                line.AppendFormat(HTML_TITLE, encounter.Title, encounter.DPS.ToString("#,##0", usCulture), combat_duration, hitpoints);

                // ---------- GROUP DAMAGE
                line.Append(HTML_TABLE_DAMAGE);
                int row = 0;
                int IsMultipleDeath = -1;
                foreach (var name in displayOrder.Values)
                {
                    string rowClass = (++row % 2 == 0) ? "odd" : "even";

                    Dictionary<string, string> entry = lines[name];
                    Dictionary<string, Dictionary<string, string>> embed;

                    if (embedLines.ContainsKey(name))
                    {
                        embed = embedLines[name];
                    }
                    else
                    {
                        embed = new Dictionary<string, Dictionary<string, string>>();
                    }

                    string nameStr = string.Format("{0}", name);

                    // a player were killed more than 1 in the report ? (-1:init / 0:no / 1: yes)
                    if (IsMultipleDeath == -1)
                    {
                        if (ALL.Equals(encounter.Title))
                        {
                            IsMultipleDeath = 1;
                        }
                        else
                        {
                            foreach (var death in entry["death"])
                            {
                                if ('0'.Equals(death) || '1'.Equals(death))
                                {
                                    continue;
                                }
                                else
                                {
                                    IsMultipleDeath = 1;
                                    break;
                                }
                            }
                            if (IsMultipleDeath == -1)
                            {
                                IsMultipleDeath = 0;
                            }
                        }
                    }
                    AppendHtmlLine(line, entry, rowClass, nameStr, false, row, IsMultipleDeath);

                    string embedRowClass = "embed" + row.ToString();
                    foreach (var embedName in embedDisplayOrder[name].Values)
                    {
                        Dictionary<string, string> embedEntry = embed[embedName];
                        string embedNameStr = embedEntry["name"];
                        AppendHtmlLine(line, embedEntry, embedRowClass, embedNameStr, true, row, 0);
                    }
                }

                // ---------- MY BUFFS
                line.Append(HTML_TABLE_BUFFS);
                if (encounter.Tags.ContainsKey(BUFFS))
                {
                    Dictionary<string, BuffData> buffs = encounter.Tags[BUFFS] as Dictionary<string, BuffData>;
                    if (buffs == null)
                    {
                        buffs = new Dictionary<string, BuffData>();
                    }

                    row = 0;
                    SortedSet<string> buffSort = new SortedSet<string>(buffs.Keys);
                    foreach (var name in buffSort)
                    {
                        var buff = buffs[name];
                        string rowClass = (++row % 2 == 0) ? "buffodd" : "buffeven";
                        double pct = buff.TotalTime * 100 / encounter.Duration.TotalMilliseconds;
                        line.AppendFormat("<tr class=\"{0}\"><td class=\"col1\"><strong>{1}</strong></td><td>{2}</td><td>{3}</td></tr>", rowClass, name, (buff.TotalTime / 1000).ToString("#,##0.0", usCulture), pct.ToString("0'%", usCulture));
                    }
                }

                // ---------- MAX HITS
                line.Append(HTML_TABLE_MAX);
                row = 0;
                foreach (var ally in encounter.GetAllies())
                {
                    var allyList = new List<CombatantData>();
                    allyList.Add(ally);

                    string maxHit;
                    string maxHeal;
                    string maxDamage;
                    GetMaxText(allyList, false, usCulture, out maxHit, out maxHeal, out maxDamage);

                    string rowClass = (++row % 2 == 0) ? "buffodd" : "buffeven";
                    line.AppendFormat("<tr class=\"{0}\"><td class=\"col1\"><strong>{1}</strong></td><td>{2}</td><td>{3}</td><td>{4}</td></tr>", rowClass, ally.Name, maxHit, maxHeal, maxDamage);
                }
                line.Append(HTML_CLOSE);

                try
                {
                    string newFile = Path.Combine(scriptFolder, "Z_act_new.html");
                    string oldFile = Path.Combine(scriptFolder, "Z_act.html");

                    using (TextWriter writer = new StreamWriter(newFile, false, Encoding.UTF8))
                    {
                        writer.WriteLine(line.ToString());
                    }

                    File.Delete(oldFile);
                    File.Move(newFile, oldFile);
                }
                catch (Exception)
                {
                    // Ignore errors writing the file.  Assume this is a security violation or such
                }

                try
                {
                    using (TextWriter writer = new StreamWriter(Path.Combine(scriptFolder, "act"), false))
                    {
                        string htmlScriptFolder = scriptFolder.Replace(Path.DirectorySeparatorChar, '/');
                        writer.WriteLine("/option WebBrowserStartURL \"file:///" + htmlScriptFolder + "/Z_act.html\"");
                        writer.WriteLine("/option web_browser 1");
                    }

                    string oldBuffFile = Path.Combine(scriptFolder, "Z_act_buffs.html");
                    if (File.Exists(oldBuffFile))
                    {
                        File.Delete(oldBuffFile);
                        File.Delete(Path.Combine(scriptFolder, "actbuffs"));
                    }
                }
                catch (Exception)
                {
                    // Ignore errors writing the file.  Assume this is a security violation or such
                }
            }
        }

        private void BuildScriptOutput(EncounterData encounter, CultureInfo usCulture, Dictionary<string, Dictionary<string, string>> lines, Dictionary<string, Dictionary<string, Dictionary<string, string>>> embedLines, SortedDictionary<string, string> displayOrder, SortedDictionary<string, SortedDictionary<string, string>> embedDisplayOrder, SortedDictionary<string, string> displayOrderHealers, SortedDictionary<string, string> displayOrderTanks)
        {
            HashSet<string> allies = new HashSet<string>();
            if (encounter.NumAllies > 0 && checkBox_ExportAllies.Checked)
            {
                foreach (var ally in encounter.GetAllies())
                {
                    allies.Add(ally.Name);
                }
            }
            bool round_dps = checkBox_ExportRoundDPS.Checked;
            foreach (var data in encounter.Items.Values)
            {
                try
                {
                    if (checkBox_Filter.Checked && checkBox_filterScript.Checked)
                    {
                        bool containsName = filterNames.Contains(data.Name);
                        if (checkBox_filterExclude.Checked)
                        {
                            if (containsName)
                            {
                                continue;
                            }
                        }
                        else
                        {
                            if (!containsName)
                            {
                                continue;
                            }
                        }
                    }

                    if (allies.Count > 0 && !allies.Contains(data.Name))
                    {
                        continue;
                    }

                    Dictionary<string, string> row = new Dictionary<string, string>();
                    row["name"] = data.Name;
                    row["death"] = data.Deaths.ToString("#,##0", usCulture);
                    row["duration"] = data.Duration.TotalSeconds > 599 ? data.DurationS : data.DurationS.Substring(1, 4);

                    row["dmg"] = data.Damage.ToString("#,##0", usCulture);
                    row["healing"] = data.Healed.ToString("#,##0", usCulture);
                    row["dmg_script"] = ConvertValue(data.Damage, usCulture);
                    row["healing_script"] = ConvertValue(data.Healed, usCulture);

                    row["dps%"] = "--".Equals(data.DamagePercent) ? "n/a" : data.DamagePercent;
                    row["hps%"] = "--".Equals(data.HealedPercent) ? "n/a" : data.HealedPercent;

                    row["dps"] = "NaN".Equals(data.DPS) ? "0" : data.DPS.ToString(round_dps?GetIntCommas():GetFloatCommas(), usCulture);

                    row["crit%"] = "NaN".Equals(data.CritDamPerc) ? "0%" : data.CritDamPerc.ToString("0'%", usCulture);
                    row["healcrit%"] = "NaN".Equals(data.CritHealPerc) ? "0%" : data.CritHealPerc.ToString("0'%", usCulture);

                    row["pen%"] = CombatantFormatSwitch(data, "Penetration%", usCulture);
                    row["glance%"] = CombatantFormatSwitch(data, "Glance%", usCulture);
                    row["block%"] = CombatantFormatSwitch(data, "Blocked%", usCulture);

                    float miss = 100.0f - data.ToHit;
                    long aegis_dmg = GetSpecialHitData(data, SecretLanguage.Aegis);
                    long aegis_heal = GetSpecialHealData(data, SecretLanguage.Aegis);
                    double aegis_dps = (data.Duration.TotalSeconds > 0) ? aegis_dmg / data.Duration.TotalSeconds : 0.0;
                    double aegis_hps = 0.0;
                    row["evade%"] = miss.ToString("0'%", usCulture);
                    row["aegisdmg"] = aegis_dmg.ToString(GetIntCommas(), usCulture);
                    row["aegisdmg_script"] = (aegis_dmg != 0)? ConvertValue(aegis_dmg, usCulture) : "";
                    row["aegisdps"] = aegis_dps.ToString(round_dps ? GetIntCommas() : GetFloatCommas(), usCulture);
                    row["aegisheal"] = aegis_heal.ToString(GetIntCommas(), usCulture);
                    row["aegisheal_script"] = (aegis_heal != 0)? ConvertValue(aegis_heal, usCulture) : "";

                    row["aegismismatch%"] = CombatantFormatSwitch(data, "AegisMismatch%", usCulture);

                    if (data.Items.ContainsKey(OUT_HEAL))
                    {
                        row["healduration"] = data.Items[OUT_HEAL].Duration.TotalSeconds > 599 ? data.Items[OUT_HEAL].DurationS : data.Items[OUT_HEAL].DurationS.Substring(1, 4);
                        row["hps"] = "NaN".Equals(data.Items[OUT_HEAL].DPS) ? "0" : data.Items[OUT_HEAL].DPS.ToString(round_dps ? GetIntCommas() : GetFloatCommas(), usCulture);
                        aegis_hps = (data.Items[OUT_HEAL].Duration.TotalSeconds > 0) ? aegis_heal / data.Duration.TotalSeconds : 0.0;
                        row["aegishps"] = aegis_hps.ToString(round_dps ? GetIntCommas() : GetFloatCommas(), usCulture);
                    }
                    else
                    {
                        row["healduration"] = "0";
                        row["hps"] = "0";
                        row["aegishps"] = "0";
                    }

                    row["takendamage"] = ConvertValue(data.DamageTaken, usCulture);
                    long aegis_inc = GetSpecialIncData(data, SecretLanguage.Aegis);
                    row["aegisinc_script"] = (aegis_inc != 0) ? ConvertValue(aegis_inc, usCulture) : "";
                    row["takencrit%"] = CombatantFormatIncSwitch(data, "takencrit%", usCulture);
                    row["takenpen%"] = CombatantFormatIncSwitch(data, "takenpen%", usCulture);
                    row["takenglance%"] = CombatantFormatIncSwitch(data, "takenglance%", usCulture);
                    row["takenblock%"] = CombatantFormatIncSwitch(data, "takenblock%", usCulture);
                    row["takenevade%"] = CombatantFormatIncSwitch(data, "takenevade%", usCulture);
                    if (data.Items.ContainsKey(INC_DAMAGE))
                    {
                        double aegis_takendps = (data.Items[INC_DAMAGE].Duration.TotalSeconds > 0) ? aegis_inc / data.Items[INC_DAMAGE].Duration.TotalSeconds : 0.0;
                        row["takenduration"] = data.Items[INC_DAMAGE].Duration.TotalSeconds > 599 ? data.Items[INC_DAMAGE].DurationS : data.Items[INC_DAMAGE].DurationS.Substring(1, 4);
                        row["takendps"] = "NaN".Equals(data.Items[INC_DAMAGE].DPS) ? "0" : data.Items[INC_DAMAGE].DPS.ToString(round_dps ? GetIntCommas() : GetFloatCommas(), usCulture);
                        row["takenaegisdps"] = aegis_takendps.ToString(round_dps ? GetIntCommas() : GetFloatCommas(), usCulture);
                    }
                    else
                    {
                        row["takenduration"] = "0";
                        row["takendps"] = "0";
                        row["takenaegisdps"] = "0";
                    }

                    string name = row["name"];
                    lines.Add(name, row);
                    displayOrder.Add(GetScriptKey(data.Damage, data.Healed, name), name);
                    displayOrderHealers.Add(GetScriptKey(data.Healed, data.Damage, name), name);
                    displayOrderTanks.Add(GetScriptKey(data.DamageTaken, data.Damage, name), name);

                    var embed = new Dictionary<string, Dictionary<string, string>>();
                    foreach (var entry in data.Items[OUT_DAMAGE].Items)
                    {
                        var item = entry.Value;
                        if (ALL.Equals(entry.Key) || item.Damage < 1)
                        {
                            continue;
                        }

                        Dictionary<string, string> att = new Dictionary<string, string>();
                        att["name"] = entry.Key;
                        att["death"] = "0";
                        att["duration"] = "";
                        att["dmg"] = item.Damage.ToString("#,##0", usCulture);
                        att["dps%"] = (item.Damage * 100 / data.Damage).ToString("0'%", usCulture);
                        att["dps"] = "0";
                        att["pen%"] = GetSpecialHitPerc(item, SecretLanguage.Penetrated).ToString("0'%", usCulture);
                        att["crit%"] = item.CritPerc.ToString("0'%", usCulture);
                        att["glance%"] = GetSpecialHitPerc(item, SecretLanguage.Glancing).ToString("0'%", usCulture);
                        att["block%"] = GetSpecialHitPerc(item, SecretLanguage.Blocked).ToString("0'%", usCulture);
                        att["healing"] = "0";
                        att["hps%"] = "0%";
                        att["hps"] = "0";
                        att["healcrit%"] = "0%";

                        miss = 100.0f - data.ToHit;
                        aegis_dmg = GetSpecialHitData(item, SecretLanguage.Aegis);
                        aegis_dps = (item.Duration.TotalSeconds > 0) ? aegis_dmg / item.Duration.TotalSeconds : 0.0;
                        att["evade%"] = miss.ToString("0'%", usCulture);
                        att["aegisdmg"] = aegis_dmg.ToString(GetIntCommas(), usCulture);
                        att["aegisdps"] = aegis_dps.ToString(round_dps ? GetIntCommas() : GetFloatCommas(), usCulture);
                        att["aegisheal"] = "0";
                        att["aegishps"] = "0.0";
                        att["aegismismatch%"] = GetSpecialHitMissPerc(item, SecretLanguage.Aegis).ToString("0'%", usCulture);

                        att["takenduration"] = "0";
                        att["takendamage"] = "0";
                        att["takendps"] = "0";
                        att["takenaegisdps"] = "0";
                        att["aegisinc_script"] = "";
                        att["takencrit%"] = "0%";
                        att["takenpen%"] = "0%";
                        att["takenglance%"] = "0%";
                        att["takenblock%"] = "0%";
                        att["takenevade%"] = "0%";

                        att["dmgKey"] = GetScriptKey(item.Damage);
                        att["healKey"] = GetScriptKey(0);
                        embed[entry.Key] = att;
                    }

                    foreach (var entry in data.Items["Outgoing Heal (Out)"].Items)
                    {
                        var item = entry.Value;
                        if (ALL.Equals(entry.Key) || item.Damage < 1)
                        {
                            continue;
                        }

                        Dictionary<string, string> att;

                        if (embed.ContainsKey(entry.Key))
                        {
                            att = embed[entry.Key];
                        }
                        else
                        {
                            att = new Dictionary<string, string>();
                            att["name"] = entry.Key;
                            att["death"] = "0";
                            att["duration"] = "";
                            att["dmg"] = "0";
                            att["dps%"] = "0%";
                            att["dps"] = "0";
                            att["pen%"] = "0%";
                            att["crit%"] = "0%";
                            att["glance%"] = "0%";
                            att["block%"] = "0%";
                            att["evade%"] = "0";
                            att["aegisdmg"] = "0";
                            att["aegisdps"] = "0.0";
                            att["aegisheal"] = "0";
                            att["aegishps"] = "0.0";
                            att["aegismismatch%"] = "0%";

                            att["takenduration"] = "0";
                            att["takendamage"] = "0";
                            att["takendps"] = "0";
                            att["takenaegisdps"] = "0";
                            att["aegisinc_script"] = "";
                            att["takencrit%"] = "0%";
                            att["takenpen%"] = "0%";
                            att["takenglance%"] = "0%";
                            att["takenblock%"] = "0%";
                            att["takenevade%"] = "0%";
                            att["dmgKey"] = GetScriptKey(0);
                        }

                        aegis_heal = GetSpecialHitData(item, SecretLanguage.Aegis);
                        aegis_hps = (item.Duration.TotalSeconds > 0) ? aegis_heal / item.Duration.TotalSeconds : 0.0;
                        att["healing"] = item.Damage.ToString("#,##0", usCulture);
                        att["hps%"] = (item.Damage * 100 / data.Healed).ToString("0'%", usCulture);
                        att["hps"] = item.EncDPS.ToString(GetFloatCommas(), usCulture);
                        att["healcrit%"] = item.CritPerc.ToString("0'%", usCulture);
                        att["aegisheal"] = aegis_heal.ToString(GetIntCommas(), usCulture);
                        att["aegishps"] = aegis_hps.ToString(GetFloatCommas(), usCulture);

                        att["healKey"] = GetScriptKey(item.Damage);

                        embed[entry.Key] = att;
                    }

                    var embedOrder = new SortedDictionary<string, string>();
                    foreach (var item in embed.Values)
                    {
                        embedOrder[GetScriptKey(item["dmgKey"], item["healKey"], item["name"])] = item["name"];
                    }

                    embedDisplayOrder[name] = embedOrder;
                    embedLines[name] = embed;
                }
                catch (Exception)
                {
                    // Ignore errors writing this entry and continue
                }
            }
        }

        #region ACTTSW
        private void SetPlayfieldInformation(string line)
        {
            string[] parts = line.Split('|');
            if (parts.Length > 1 && line.Length > 21)
            {
                try
                {
                    string zoneName = parts[1].Trim();
                    if (zoneName.Length < 1)
                    {
                        zoneName = "Unknown Zone";
                    }

                    lock (locker)
                    {
                        if (!ActGlobals.oFormActMain.CurrentZone.Equals(zoneName))
                        {
                            ActGlobals.oFormActMain.BeginInvoke(new MethodInvoker(delegate
                            {
                                ActGlobals.oFormActMain.ChangeZone(zoneName);
                            }));
                        }
                    }
                }
                catch
                {
                }
            }
        }

        private void SetEvadeInformation(string line)
        {
            string[] parts = line.Split('|');
            if (parts.Length > 2 && line.Length > 21)
            {
                try
                {
                    string[] dateParts = parts[0].Split(' ');
                    DateTime detectedTime = XmlConvert.ToDateTime(dateParts[0].TrimStart('[') + "T" + dateParts[1], XmlDateTimeSerializationMode.Local);
                    if (!detectedTime.Equals(lastEvadeTime))
                    {
                        SetEvadeAction(parts[1], parts[2], detectedTime);
                        lastEvadeTime = detectedTime;
                    }
                }
                catch
                {
                }
            }
        }

        private void SetEvadeAction(string attacker, string victim, DateTime detectedTime)
        {
            lock (locker)
            {
                if (ActGlobals.oFormActMain.InCombat)
                {
                    if (ActGlobals.oFormActMain.LastKnownTime != DateTime.MinValue)
                    {
                        TimeSpan timeDiff = detectedTime.Subtract(ActGlobals.oFormActMain.LastKnownTime);
                        if (timeDiff.TotalSeconds < 0 || timeDiff.TotalSeconds > 6)
                        {
                            detectedTime = ActGlobals.oFormActMain.LastKnownTime;
                        }
                    }

                    ActGlobals.oFormActMain.BeginInvoke(new MethodInvoker(delegate
                    {
                        if (ActGlobals.oFormActMain.SetEncounter(detectedTime, attacker, victim))
                        {
                            ActGlobals.oFormActMain.AddCombatAction((int)SwingTypeEnum.Melee, false, "", attacker, "Evade", 0, detectedTime, ActGlobals.oFormActMain.GlobalTimeSorter, victim, SecretLanguage.none);
                        }
                    }));
                }
            }
        }

        private void ResetCharName(string line)
        {
            string[] parts = line.Split('|');
            if (parts.Length > 1)
            {
                lock (locker)
                {
                    SetCharName(parts[1]);
                }
            }
        }

        private void SetStartBuffs(string line)
        {
            string[] parts = line.Split('|');
            if (parts.Length > 2)
            {
                string[] buffs = parts[2].Split(':');
                HashSet<string> buffNames = new HashSet<string>();
                for (int indx = 1; indx < buffs.Length; ++indx)
                {
                    if (!buffNames.Contains(buffs[indx]))
                    {
                        buffNames.Add(buffs[indx]);
                    }
                }

                lock (locker)
                {
                    List<string> toDelete = new List<string>();
                    foreach (var buff in buffData)
                    {
                        if (!buffNames.Contains(buff.Key) && !buff.Value.Changed)
                        {
                            toDelete.Add(buff.Key);
                        }
                    }

                    for (int indx = 0; indx < toDelete.Count; ++indx)
                    {
                        buffData.Remove(toDelete[indx]);
                    }
                }
            }
        }

        private void LeaveCombat(string line)
        {
            string[] parts = line.Split('|');
            if (parts.Length > 1)
            {
                lock (locker)
                {
                    SetCharName(parts[1]);

                    allies = new HashSet<string>();
                    for (int i = 1; i < parts.Length; ++i)
                    {
                        if (parts[i].Length > 0 && !allies.Contains(parts[i]))
                        {
                            allies.Add(parts[i]);
                        }
                    }

                    if (ActGlobals.oFormActMain.InCombat)
                    {
                        ActGlobals.oFormActMain.BeginInvoke(new MethodInvoker(delegate { ActGlobals.oFormActMain.EndCombat(true); }));
                    }
                }
            }
        }

        private void EndOldCombat()
        {
            lock (locker)
            {
                if (ActGlobals.oFormActMain.InCombat && lastCombatStart != DateTime.MinValue)
                {
                    TimeSpan diff = ActGlobals.oFormActMain.LastKnownTime.Subtract(lastCombatStart);
                    if (diff.TotalSeconds > 3)
                    {
                        ActGlobals.oFormActMain.BeginInvoke(new MethodInvoker(delegate { ActGlobals.oFormActMain.EndCombat(true); }));
                    }
                }
            }
        }

        private static void CheckTSWACTEnabled(SecretParse parser)
        {
            try
            {
                while (runClientReadThread == 1)
                {
                    bool readFile = false;

                    lock (parser.addonEnabledLocker)
                    {
                        readFile = parser.addonEnabled;
                    }

                    if (readFile)
                    {
                        parser.ReadClientFile();
                    }
                    else
                    {
                        Thread.Sleep(5000);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private void ReadClientFile()
        {
            string filePath = Path.Combine(Path.GetDirectoryName(ActGlobals.oFormActMain.LogFilePath), "ClientLog.txt");
            long offset = 0;

            if (!File.Exists(filePath))
            {
                using (File.CreateText(filePath))
                {
                }
            }

            using (FileStream file = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Read to end of file looking for character name
                using (StreamReader reader = new StreamReader(file))
                {
                    if (!reader.EndOfStream)
                    {
                        do
                        {
                            ProcessClientLogLine(reader.ReadLine(), true);
                        } while (!reader.EndOfStream);

                        offset = file.Position;
                    }
                }
            }

            using (FileStream file = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (StreamReader reader = new StreamReader(file))
                {
                    while (runClientReadThread == 1)
                    {
                        file.Seek(offset, SeekOrigin.Begin);
                        if (!reader.EndOfStream)
                        {
                            do
                            {
                                ProcessClientLogLine(reader.ReadLine(), false);
                            } while (!reader.EndOfStream);

                            offset = file.Position;
                        }

                        if (combatEnding != null)
                        {
                            LeaveCombat(combatEnding);
                            combatEnding = null;
                        }

                        Thread.Sleep(500);

                        lock (addonEnabledLocker)
                        {
                            if (!addonEnabled)
                            {
                                return;
                            }
                        }
                    }
                }
            }
        }

        private void ProcessClientLogLine(string inLine, bool startup)
        {
            if (inLine.Contains("TSWACT") && inLine.Contains("|"))
            {
                if (inLine.Contains("Loaded"))
                {
                    ResetCharName(inLine);
                }
                else if (inLine.Contains("Enter combat"))
                {
                    EndOldCombat();
                    ResetCharName(inLine);
                    SetStartBuffs(inLine);
                    combatEnding = null;
                }
                else if (inLine.Contains("- Evade -"))
                {
                    if (!startup)
                    {
                        SetEvadeInformation(inLine);
                    }
                }
                else if (inLine.Contains("- Playfield -"))
                {
                    SetPlayfieldInformation(inLine);
                }
                else if (inLine.Contains("Out of combat"))
                {
                    if (startup)
                    {
                        ResetCharName(inLine);
                    }
                    else
                    {
                        if (combatEnding == null)
                        {
                            // Sometimes we have a few late hits, so give them a chance to arrive
                            Thread.Sleep(1250);
                            combatEnding = inLine;
                        }
                        else
                        {
                            LeaveCombat(inLine);
                            combatEnding = null;
                        }
                    }
                }
            }
        }

        private void FixDBConfFile()
        {
            if (!checkBox_AutofixDBConf.Checked || !checkBox_EnableTSWAddon.Checked)
            {
                return;
            }

            try
            {
                bool backupCurrentFile = true;
                string backupFilePath = Path.Combine(Path.GetDirectoryName(ActGlobals.oFormActMain.LogFilePath), "dbDebug.conf.actold.txt");
                string backupFilePath2 = Path.Combine(Path.GetDirectoryName(ActGlobals.oFormActMain.LogFilePath), "dbDebug.conf.actolder.txt");
                if (File.Exists(backupFilePath2))
                {
                    if (File.Exists(backupFilePath))
                    {
                        File.Delete(backupFilePath2);
                    }
                    else
                    {
                        File.Move(backupFilePath2, backupFilePath);
                    }
                }

                List<string> lines = new List<string>();
                string filePath = Path.Combine(Path.GetDirectoryName(ActGlobals.oFormActMain.LogFilePath), "dbDebug.conf");
                ReadConfFile(lines, filePath);
                if (lines.Count == 0)
                {
                    ReadConfFile(lines, backupFilePath);
                    backupCurrentFile = false;
                }

                bool changed = false;
                List<string> newLines = new List<string>();
                if (lines.Count > 10)
                {
                    foreach (var line in lines)
                    {
                        if (line.StartsWith("DebugLevel="))
                        {
                            string tempLine = "DebugLevel=4";
                            if (tempLine.Equals(line.Trim()))
                            {
                                newLines.Add(line);
                            }
                            else
                            {
                                changed = true;
                                newLines.Add(tempLine);
                            }
                        }
                        else if (line.StartsWith("PrintModule="))
                        {
                            string tempLine = "PrintModule=true";
                            if (tempLine.Equals(line.Trim()))
                            {
                                newLines.Add(line);
                            }
                            else
                            {
                                changed = true;
                                newLines.Add(tempLine);
                            }
                        }
                        else if (line.StartsWith("LogFile=Patcher.log"))
                        {
                            // The patcher is currently running, so we don't want to touch this file
                            return;
                        }
                        else
                        {
                            newLines.Add(line);
                        }
                    }
                }

                if (changed)
                {
                    string newFilePath = Path.Combine(Path.GetDirectoryName(ActGlobals.oFormActMain.LogFilePath), "dbDebug.conf.act.txt");
                    File.Delete(newFilePath);
                    using (StreamWriter writer = new StreamWriter(newFilePath))
                    {
                        foreach (var line in newLines)
                        {
                            writer.WriteLine(line);
                        }
                    }

                    if (backupCurrentFile)
                    {
                        if (File.Exists(backupFilePath))
                        {
                            File.Delete(backupFilePath2);
                            File.Move(backupFilePath, backupFilePath2);
                        }

                        File.Move(filePath, backupFilePath);
                        File.Move(newFilePath, filePath);
                        File.Delete(backupFilePath2);
                    }
                    else
                    {
                        File.Delete(filePath);
                        File.Move(newFilePath, filePath);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void ReadConfFile(List<string> lines, string filePath)
        {
            if (File.Exists(filePath))
            {
                using (FileStream file = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (StreamReader reader = new StreamReader(file))
                    {
                        while (!reader.EndOfStream)
                        {
                            lines.Add(reader.ReadLine());
                        }
                    }
                }
            }
        }
        #endregion

        private void SetCharName(string newName)
        {
            charName = newName;
            if (charName == null)
            {
                if (ActGlobals.charName != null)
                {
                    charName = ActGlobals.charName;
                }
                else
                {
                    charName = "";
                }
            }
        }

        private string GetCharName()
        {
            if (charName != null && charName.Length > 0)
            {
                return charName;
            }

            if (ActGlobals.charName != null && ActGlobals.charName.Length > 0)
            {
                return ActGlobals.charName;
            }

            return "";
        }

        private string ConvertCharName(string inName)
        {
            string tempCharName = GetCharName();

            if (inName.Length < 1 && SecretLanguage.Language.Equals(SecretLanguage.German))
            {
                inName = SecretLanguage.You;
            }

            string lowerName = inName.ToLowerInvariant();
            if (SecretLanguage.YouSet.Contains(lowerName))
            {
                if (tempCharName != null && tempCharName.Length > 0)
                {
                    return tempCharName;
                }
                else
                {
                    return SecretLanguage.You.ToUpper();
                }
            }

            if (inName.EndsWith("'s"))
            {
                inName = inName.Substring(0, inName.Length - 2);
            }

            if (tempCharName != null && tempCharName.Equals(inName))
            {
                return tempCharName;
            }

            if (inName.Length < 1 && !checkBox_HideUnknown.Checked)
            {
                inName = "TSW_Unknown";
            }

            return inName;
        }

        private string TrimBrackets(string inStr)
        {
            return inStr.Trim(' ', '(', ')');
        }

        private string AddSpecial(string special, string addition)
        {
            if (SecretLanguage.none.Equals(special))
            {
                return addition;
            }

            return special + "," + addition;
        }

        private string removeTimestamp(string inStr)
        {
            return timeStamp.Replace(inStr, "");
        }

        private string removeFont(string inStr)
        {
            return fontHtml.Replace(inStr, "");
        }

        private string ConvertValue(long value, CultureInfo usCulture)
        {
            string[] ext = { "", "k", "m", "g" };
            int idx = 0;
            float calc = value;
            while (calc > 500 && idx < ext.Length-1)
            {
                calc /= 1000;
                idx++;
            }
            return calc.ToString((idx==0)?"0":"0.00", usCulture) + ext[idx];
        }

        void oFormActMain_BeforeLogLineRead(bool isImport, LogLineEventArgs logInfo)
        {
            try
            {
                lock (locker)
                {
                    string tempInputStr = previousLine + logInfo.logLine;
                    string oldPreviousLine = previousLine;
                    previousLine = "";

                    string InputStr = removeTimestamp(tempInputStr);
                    int pos = InputStr.IndexOf('[');
                    if (pos > 0)
                    {
                        previousLine = InputStr.Substring(pos);
                        InputStr = InputStr.Substring(0, pos);
                    }

                    string Secret_attacker = string.Empty;
                    string victim = string.Empty;
                    string attacker = string.Empty;
                    string damage = string.Empty;
                    int attackType = (int)SwingTypeEnum.Melee;
                    string effectName = string.Empty;
                    string attackName = string.Empty;
                    string special = SecretLanguage.none;
                    string Aowner = string.Empty;
                    string Vowner = string.Empty;
                    Boolean SelfAttack = false;
                    Int64 Amount = 0;

                    Boolean critical = false;
                    Boolean colorlog = true;
                    string attackSuffix = "";

                    int eventType = -1;

                    MatchCollection matches = null;

                    if (InputStr.IndexOf(SecretLanguage.HitLine1) != -1 || InputStr.IndexOf(SecretLanguage.HitLine2) != -1 || InputStr.IndexOf(SecretLanguage.HateLine) != -1)
                    {
                        foreach (var damageLine in SecretLanguage.damageLines)
                        {
                            matches = damageLine.Matches(InputStr);
                            if (matches != null && matches.Count > 0)
                            {
                                break;
                            }
                        }

                        if (matches != null && matches.Count == 1)
                        {
                            eventType = 3;
                            ++dpsCount;
                        }
                        else
                        {
                            eventType = -1;
                        }
                    }

                    if (InputStr.IndexOf(SecretLanguage.RedirectedLine) != -1)
                    {
                        foreach (var redirectLine in SecretLanguage.redirectLines)
                        {
                            matches = redirectLine.Matches(InputStr);

                            if (matches != null && matches.Count > 0)
                            {
                                break;
                            }
                        }

                        if (matches != null && matches.Count == 1)
                        {
                            attackSuffix = " (redirect)";
                            eventType = 3;
                            ++dpsCount;
                        }
                        else
                        {
                            eventType = -1;
                        }
                    }

                    if (eventType == -1 && (InputStr.IndexOf(SecretLanguage.HealLine1) != -1 || InputStr.IndexOf(SecretLanguage.HealLine2) != -1 || InputStr.IndexOf(SecretLanguage.AbsorbLine) != -1))
                    {
                        foreach (var healLine in SecretLanguage.healLines)
                        {
                            matches = healLine.Matches(InputStr);

                            if (matches != null && matches.Count > 0)
                            {
                                break;
                            }
                        }

                        if (matches != null && matches.Count == 1)
                        {
                            eventType = 5;
                            ++healCount;

                            if (InputStr.IndexOf(SecretLanguage.AbsorbLine) != -1)
                            {
                                attackSuffix = " (absorb)";
                            }
                        }
                        else
                        {
                            eventType = -1;
                        }
                    }

                    if (eventType == -1 && (InputStr.IndexOf(SecretLanguage.EvadeLine1) != -1 || InputStr.IndexOf(SecretLanguage.EvadeLine2) != -1))
                    {
                        foreach (var evadedLine in SecretLanguage.evadedLines)
                        {
                            matches = evadedLine.Matches(InputStr);

                            if (matches != null && matches.Count > 0)
                            {
                                break;
                            }
                        }

                        if (matches != null && matches.Count == 1)
                        {
                            eventType = 10;
                        }
                        else
                        {
                            eventType = -1;
                        }
                    }

                    if (eventType == -1 && InputStr.IndexOf(SecretLanguage.DiedLine) != -1)
                    {
                        foreach (var diedLine in SecretLanguage.diedLines)
                        {
                            matches = diedLine.Matches(InputStr);

                            if (matches != null && matches.Count > 0)
                            {
                                break;
                            }
                        }

                        if (matches != null && matches.Count == 1)
                        {
                            eventType = 12;
                        }
                        else
                        {
                            eventType = -1;
                        }
                    }

                    if (eventType == -1 && (InputStr.IndexOf(SecretLanguage.BuffLine1) != -1 || InputStr.IndexOf(SecretLanguage.BuffLine2) != -1))
                    {
                        foreach (var buffLine in SecretLanguage.buffLines)
                        {
                            matches = buffLine.Matches(InputStr);

                            if (matches != null && matches.Count > 0)
                            {
                                GroupCollection groups = matches[0].Groups;
                                string name = groups["name"].Value.Trim(' ', '"', '\'', '.');
                                if (!buffData.ContainsKey(name))
                                {
                                    buffData.Add(name, new BuffData());
                                }

                                buffData[name].Start(logInfo.detectedTime);
                                break;
                            }
                        }

                        eventType = -1;
                    }

                    if (eventType == -1 && (InputStr.IndexOf(SecretLanguage.BuffStopLine) != -1))
                    {
                        foreach (var buffLine in SecretLanguage.buffStopLines)
                        {
                            matches = buffLine.Matches(InputStr);

                            if (matches != null && matches.Count > 0)
                            {
                                GroupCollection groups = matches[0].Groups;
                                string name = groups["name"].Value.Trim(' ', '"');
                                if (buffData.ContainsKey(name))
                                {
                                    buffData[name].Stop(logInfo.detectedTime);
                                }

                                break;
                            }
                        }

                        eventType = -1;
                    }
                    ProcessLogLineEntry(logInfo, ref victim, ref attacker, ref attackType, ref attackName, ref special, ref SelfAttack, ref Amount, ref critical, colorlog, attackSuffix, eventType, matches);
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message + "\n" + e.StackTrace);
            }
        }

        private void ProcessLogLineEntry(LogLineEventArgs logInfo, ref string victim, ref string attacker, ref int attackType, ref string attackName, ref string special, ref Boolean SelfAttack, ref Int64 Amount, ref Boolean critical, Boolean colorlog, string attackSuffix, int eventType, MatchCollection matches)
        {
            if (eventType != -1) // Filtering
            {

                ActGlobals.oFormActMain.GlobalTimeSorter++;

                GroupCollection groups = matches[0].Groups;
                attacker = groups["actor"].Value.Trim(' ', '"');
                victim = groups["actee"].Value.Trim(' ', '"');
                attackName = removeFont(groups["attackName"].Value);
                if (("".Equals(attacker)) && !(attackName.Contains("\"")) && (SecretLanguage.Language == SecretLanguage.German))
                {
                    attacker = victim;
                }
                attackName = attackName.Trim(' ', '"', '\'') + attackSuffix;
                string damageTypeStr = TrimBrackets(groups["damageType"].Value);
                string damageClassStr = TrimBrackets(groups["damageClass"].Value);
                string blockTypeStr = TrimBrackets(groups["blockType"].Value);
                string amountStr = TrimBrackets(groups["amount"].Value);

                try
                {
                    if (amountStr == "")
                    {
                        Amount = 0;
                    }
                    else
                    {
                        Amount = Int64.Parse(amountStr);
                    }
                }
                catch (Exception)
                {
                    Amount = 0;
                }

                critical = false;
                if (groups["crit"].Length > 0)
                {
                    critical = true;
                }

                attacker = ConvertCharName(attacker);

                if ("Ihres".Equals(victim))
                {
                    victim = attacker;
                }
                else
                {
                    victim = ConvertCharName(victim);
                }

                // Ignore unknown attacks
                if (eventType == 12)
                {
                    if (attacker.Length < 1)
                    {
                        return;
                    }
                }
                else
                {
                    if (attacker.Length < 1 || victim.Length < 1)
                    {
                        return;
                    }
                }

                // Filter the results if enabled
                if (filterNames.Count > 0 && checkBox_Filter.Checked)
                {
                    if (checkBox_filterExclude.Checked)
                    {
                        if (filterNames.Contains(attacker) || filterNames.Contains(victim))
                        {
                            return;
                        }
                    }
                    else
                    {
                        if (!(filterNames.Contains(attacker) || filterNames.Contains(victim)))
                        {
                            return;
                        }
                    }
                }

                string DamageType;
                if (damageClassStr.Length > 0)
                {
                    DamageType = damageClassStr.Trim();
                }
                else
                {
                    DamageType = SecretLanguage.none;
                }

                if (SecretLanguage.Glancing.Equals(damageTypeStr))
                {
                    special = AddSpecial(special, damageTypeStr);
                }

                if (SecretLanguage.Penetrated.Equals(blockTypeStr))
                {
                    special = AddSpecial(special, blockTypeStr);
                }

                if (SecretLanguage.Blocked.Equals(blockTypeStr))
                {
                    special = AddSpecial(special, blockTypeStr);
                }

                // AEGIS
                if (victim.IndexOf(SecretLanguage.AegisShieldLine) != -1)
                {
                    foreach (var aegisShieldLine in SecretLanguage.aegisShieldLines)
                    {
                        matches = aegisShieldLine.Matches(victim);
                        if (matches != null && matches.Count > 0)
                        {
                            GroupCollection aegisGroup = matches[0].Groups;
                            string actee = aegisGroup["actee"].Value.Trim(' ', '"', '\'', '.');
                            string aegis = aegisGroup["aegis"].Value.Trim(' ', '"', '\'', '.');

                            if ((actee.Length > 0) && (aegis.Length > 0))
                            {
                                special = AddSpecial(special, SecretLanguage.Aegis);
                                if (DamageType == SecretLanguage.none)
                                {
                                    DamageType = (Amount > 0) ? aegis : "Aegis Mismatch";
                                }
                                actee = ConvertCharName(actee);
                                victim = (checkBox_ReduceAegis.Checked)? actee : actee + " - " + aegis + " AEGIS";
                                break;
                            }
                        }
                    }
                }

                // Ignore self hits
                if (attacker.Equals(victim))
                {
                    SelfAttack = true;
                }


                #region Main Parsing coding
                // Parsing codes
                //
                switch (eventType)
                {
                    #region Case 3 [Hit]
                    case 3:
                        // Normal attack
                        // <Attacker> , <Vicitim> , <Amount> , <Spell ID> , <Spell Name> )
                        // <attacker>'s <spell name> hits <Victim> for <Amount> damage.
                        if (Amount < 0) Amount = 0;

                        if ((((SelfAttack && ActGlobals.oFormActMain.InCombat) || !SelfAttack) && "" != attacker && "" != victim && ActGlobals.oFormActMain.SetEncounter(logInfo.detectedTime, attacker, victim)) || (ActGlobals.oFormActMain.InCombat && ("" == attacker || "" == victim))) // Altuslumen 1.3.0.3.2 one line edit
                        {
                            if (SelfAttack)
                            {
                                if (checkBox_SelfDamage.Checked)
                                {
                                    string newAttacker = "";
                                    if (checkBox_SelfPlayerDamage.Checked)
                                    {
                                        newAttacker = attacker;
                                    }
                                    newAttacker += "_self_";
                                    ActGlobals.oFormActMain.AddCombatAction(attackType, critical, special, newAttacker, attackName, Amount, logInfo.detectedTime, ActGlobals.oFormActMain.GlobalTimeSorter, victim, DamageType);
                                }
                                else
                                {
                                    ActGlobals.oFormActMain.AddCombatAction(attackType, critical, special, attacker, attackName, new Dnum(0, "0 (" + Amount + ")"), logInfo.detectedTime, ActGlobals.oFormActMain.GlobalTimeSorter, victim, DamageType);
                                }
                                //ActGlobals.oFormActMain.AddDamageToGraph(attacker, 0);
                            }
                            else
                            {
                                ActGlobals.oFormActMain.AddCombatAction(attackType, critical, special, attacker, attackName, Amount, logInfo.detectedTime, ActGlobals.oFormActMain.GlobalTimeSorter, victim, DamageType);
                                //ActGlobals.oFormActMain.AddDamageToGraph(attacker, Amount);
                            }

                            if (colorlog) logInfo.detectedType = System.Drawing.Color.Red.ToArgb();
                        }
                        break;
                    #endregion

                    #region Case 5 [Heals]
                    case 5:
                        // Someone is healed
                        // <Attacker> , <Vicitim> , <Amount> , <Spell ID> , <Spell Name> )
                        // <Attacker>'s <Spell Name> heals <Vicitim> for <Amount>.

                        // Check if overheal exists
                        Dnum DamN = new Dnum(Amount);
                        attackType = (int)SwingTypeEnum.Healing;

                        if (ActGlobals.oFormActMain.InCombat)
                        {
                            ActGlobals.oFormActMain.AddCombatAction(attackType, critical, special, attacker, attackName, DamN, logInfo.detectedTime, ActGlobals.oFormActMain.GlobalTimeSorter, victim, DamageType);
                            if (colorlog) logInfo.detectedType = System.Drawing.Color.Green.ToArgb();
                        }
                        break;
                    #endregion

                    #region Case 10 [Miss]
                    case 10:
                        // An attack misses
                        // 10 , T=N#R=O#<ID Number of attacker> , T=P#R=O#<ID Number of Victim> , T=X#R=X#<ID Number of attacker's owner if pet> , T=X#R=X#<ID Number of attacker's owner if pet> ,
                        // <Attacker> , <Vicitim> , <Amount> , <Spell ID> , <Spell Name> )
                        // <Attacker>'s <Spell Name> misses <Vicitim>.
                        if (ActGlobals.oFormActMain.SetEncounter(logInfo.detectedTime, attacker, victim))
                        {
                            ActGlobals.oFormActMain.AddCombatAction(attackType, critical, special, attacker, attackName, new Dnum(Dnum.Miss, SecretLanguage.Miss), logInfo.detectedTime, ActGlobals.oFormActMain.GlobalTimeSorter, victim, DamageType); // Altuslumen 1.3.0.3.2
                            if (colorlog) logInfo.detectedType = System.Drawing.Color.Blue.ToArgb();
                        }
                        break;
                    #endregion

                    #region Case 12 [Death]
                    case 12:
                        // Someone died
                        // 12 , T=N#R=O#<ID Number of attacker> , T=P#R=O#<ID Number of Victim> , T=X#R=X#<ID Number of attacker's owner if pet> , T=X#R=X#<ID Number of attacker's owner if pet> ,
                        // <Attacker> , <Vicitim> , <Amount> , <Spell ID> , <Spell Name> )

                        if (ActGlobals.oFormActMain.InCombat)
                        {
                            ActGlobals.oFormActMain.AddCombatAction(attackType, critical, SecretLanguage.none, attacker, SecretLanguage.Killing, Dnum.Death, logInfo.detectedTime, ActGlobals.oFormActMain.GlobalTimeSorter, attacker, DamageType); // Altuslumen 1.3.0.3.2
                            if (colorlog) logInfo.detectedType = System.Drawing.Color.Red.ToArgb();
                            //if (ActGlobals.oFormActMain.CbKillEnd_Checked && ActGlobals.oFormActMain.ActiveZone.ActiveEncounter.GetAllies().Contains(new CombatantData(CodeList[5], null))) ActGlobals.oFormActMain.EndCombat(true);
                        }
                        break;
                    #endregion

                }
                #endregion
            }
        }

        private DateTime ParseDateTime(string FullLogLine)
        {
            // Thanks Tharangus for modification to adjust the timezones in the logs.
            string timeStr = FullLogLine.Substring(1, 8);
            DateTime val = DateTime.ParseExact(timeStr, "HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
            if (val.CompareTo(ActGlobals.oFormActMain.LastKnownTime) < 0)
            {
                val = val.AddDays(1);
            }

            return val;
        }


        void LoadSettings()
        {
            #region Control Settings

            xmlSettings.AddControlSetting(comboBox_Language.Name, comboBox_Language);
            xmlSettings.AddControlSetting(checkBox_AutoCheck.Name, checkBox_AutoCheck);
            xmlSettings.AddControlSetting(checkBox_ExportScript.Name, checkBox_ExportScript);
            xmlSettings.AddControlSetting(checkBox_LimitNames.Name, checkBox_LimitNames);
            xmlSettings.AddControlSetting(checkBox_ExportRoundDPS.Name, checkBox_ExportRoundDPS);
            xmlSettings.AddControlSetting(checkBox_ExportShowLegend.Name, checkBox_ExportShowLegend);
            xmlSettings.AddControlSetting(checkBox_ExportColored.Name, checkBox_ExportColored);
            xmlSettings.AddControlSetting(checkBox_ExportSplit.Name, checkBox_ExportSplit);
            xmlSettings.AddControlSetting(checkBox_ExportHtml.Name, checkBox_ExportHtml);
            xmlSettings.AddControlSetting(checkBox_ExportAllies.Name, checkBox_ExportAllies);
            xmlSettings.AddControlSetting(checkBox_DontExportShortEnc.Name, checkBox_DontExportShortEnc);
            xmlSettings.AddControlSetting(checkBox_EnableTSWAddon.Name, checkBox_EnableTSWAddon);
            xmlSettings.AddControlSetting(checkBox_AutofixDBConf.Name, checkBox_AutofixDBConf);
            xmlSettings.AddControlSetting(checkBox_HideUnknown.Name, checkBox_HideUnknown);
            xmlSettings.AddControlSetting(checkBox_SelfDamage.Name, checkBox_SelfDamage);
            xmlSettings.AddControlSetting(checkBox_SelfPlayerDamage.Name, checkBox_SelfPlayerDamage);
            xmlSettings.AddControlSetting(checkBox_ReduceAegis.Name, checkBox_ReduceAegis);
            xmlSettings.AddControlSetting(checkBox_Filter.Name, checkBox_Filter);
            xmlSettings.AddControlSetting(checkBox_filterExclude.Name, checkBox_filterExclude);
            xmlSettings.AddControlSetting(checkBox_filterScript.Name, checkBox_filterScript);
            xmlSettings.AddControlSetting(checkedListBox_ExportFields.Name, checkedListBox_ExportFields);
            xmlSettings.AddControlSetting(checkedListBox_Filters.Name, checkedListBox_Filters);
            #endregion

            if (File.Exists(settingsFile))
            {
                FileStream fs = new FileStream(settingsFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                XmlTextReader xReader = new XmlTextReader(fs);

                try
                {
                    while (xReader.Read())
                    {
                        if (xReader.NodeType == XmlNodeType.Element)
                        {
                            if (xReader.LocalName == "SettingsSerializer")
                            {
                                xmlSettings.ImportFromXml(xReader);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    lblStatus.Text = "Error loading settings: " + ex.Message;
                }
                xReader.Close();
            }

            SetFilterEnabled(checkBox_Filter.Checked);
            SetLanguage(comboBox_Language.Text);
        }

        void SaveSettings()
        {
            FileStream fs = new FileStream(settingsFile, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
            XmlTextWriter xWriter = new XmlTextWriter(fs, Encoding.UTF8);
            xWriter.Formatting = Formatting.Indented;
            xWriter.Indentation = 1;
            xWriter.IndentChar = '\t';
            xWriter.WriteStartDocument(true);
            xWriter.WriteStartElement("Config");	// <Config>
            xWriter.WriteStartElement("SettingsSerializer");	// <Config><SettingsSerializer>
            xmlSettings.ExportToXml(xWriter);	// Fill the SettingsSerializer XML
            xWriter.WriteEndElement();	// </SettingsSerializer>
            xWriter.WriteEndElement();	// </Config>
            xWriter.WriteEndDocument();	// Tie up loose ends (shouldn't be any)
            xWriter.Flush();	// Flush the file buffer to disk
            xWriter.Close();
        }

        void SecretCheckUpdate()
        {
            if (!checkBox_AutoCheck.Checked) return;
            int pluginId = 63; // The download ID from the ACT website
            try
            {
                DateTime localDate = ActGlobals.oFormActMain.PluginGetSelfDateUtc(this);
                DateTime remoteDate = ActGlobals.oFormActMain.PluginGetRemoteDateUtc(pluginId);
                if (localDate.AddHours(2) < remoteDate)	// This gives you a couple hours to zip+upload the plugin since the last file change
                {
                    DialogResult ResultBox = MessageBox.Show("There is a new version of the Secret Parsing Plugin available. \n" +
                                                             "You may download this from http://advancedcombattracker.com/download.php.\n\n" +
                                                             "Do you want to try auto download and install the plugin?",
                                                             "Secret Plugin Version Check", MessageBoxButtons.YesNo);
                    if (ResultBox == DialogResult.Yes)
                    {
                        FileInfo updatedFile = ActGlobals.oFormActMain.PluginDownload(pluginId);
                        ActPluginData pluginData = ActGlobals.oFormActMain.PluginGetSelfData(this);

                        // If the remote file is a .cs/.vb/.dll file
                        pluginData.pluginFile.Delete();
                        updatedFile.MoveTo(pluginData.pluginFile.FullName);

                        ThreadInvokes.CheckboxSetChecked(ActGlobals.oFormActMain, pluginData.cbEnabled, false);
                        Application.DoEvents();
                        ThreadInvokes.CheckboxSetChecked(ActGlobals.oFormActMain, pluginData.cbEnabled, true);
                    }
                }

                SecretCheckGitUpdate();
            }
            catch
            {
            }
        }

        void SecretCheckGitUpdate()
        {
            string gitUrl = "https://github.com/Inkraja/Advanced-Combat-Tracker/releases/latest";
            string gitFile = "https://raw.githubusercontent.com/Inkraja/Advanced-Combat-Tracker/{0}/SecretParser.cs";

            System.Net.WebRequest request = System.Net.WebRequest.Create(gitUrl);
            System.Net.WebResponse response = request.GetResponse();

            try
            {
                if (response.ResponseUri.Segments.Length > 0)
                {
                    string newVersion = response.ResponseUri.Segments[response.ResponseUri.Segments.Length - 1];
                    System.Version ver = Assembly.GetExecutingAssembly().GetName().Version;
                    string curVersion = ver.ToString();

                    if (newVersion.CompareTo(curVersion) > 0)
                    {
                        DialogResult ResultBox = MessageBox.Show("There is a new version " + newVersion + " of the Secret Parsing Plugin (currently " + curVersion + ") available. \n" +
                                                                 "You may download this from https://github.com/Inkraja/Advanced-Combat-Tracker/releases/latest.\n\n" +
                                                                 "Do you want to try auto download and install the plugin?",
                                                                 "Secret Plugin Version Check", MessageBoxButtons.YesNo);
                        if (ResultBox == DialogResult.Yes)
                        {
                            System.Net.WebRequest reqFile = System.Net.WebRequest.Create(String.Format(gitFile, newVersion));
                            System.Net.WebResponse resFile = reqFile.GetResponse();

                            if (resFile != null)
                            {
                                ActPluginData pluginData = ActGlobals.oFormActMain.PluginGetSelfData(this);
                                string localFile = Path.GetTempPath() + pluginData.pluginFile.Name;

                                Stream remoteStream = resFile.GetResponseStream();
                                Stream localStream = File.Create(localFile);
                                int bytesTotal = 0;

                                try
                                {
                                    byte[] buffer = new byte[1024];
                                    int bytesRead;

                                    do
                                    {
                                        bytesRead = remoteStream.Read(buffer, 0, buffer.Length);
                                        localStream.Write(buffer, 0, bytesRead);
                                        bytesTotal += bytesRead;
                                    } while (bytesRead > 0);
                                }
                                finally
                                {
                                    resFile.Close();
                                    remoteStream.Close();
                                    localStream.Close();
                                }

                                if (bytesTotal > 0)
                                {
                                    FileInfo updatedFile = new FileInfo(localFile);
                                    pluginData.pluginFile.Delete();
                                    updatedFile.MoveTo(pluginData.pluginFile.FullName);

                                    ThreadInvokes.CheckboxSetChecked(ActGlobals.oFormActMain, pluginData.cbEnabled, false);
                                    Application.DoEvents();
                                    ThreadInvokes.CheckboxSetChecked(ActGlobals.oFormActMain, pluginData.cbEnabled, true);
                                }

                            }

                        }
                    }

                }
            }
            finally
            {
                response.Close();
            }
        }

        private void SetFilterEnabled(bool enabled)
        {
            textBox_FilterName.Enabled = enabled;
            textBox_FilterName.Text = "";
            button_AddCharacter.Enabled = enabled;
            button_DeleteCharacter.Enabled = enabled;
            checkedListBox_Filters.Enabled = enabled;
            checkBox_filterExclude.Enabled = enabled;
            checkBox_filterScript.Enabled = enabled;

            filterNames.Clear();
            if (enabled)
            {
                foreach (var objName in checkedListBox_Filters.CheckedItems)
                {
                    string name = (string)objName;
                    if (!filterNames.Contains(name))
                    {
                        filterNames.Add(name);
                    }
                }
            }
        }

        private void AddFilterName(string inName)
        {
            string name = inName.Trim();

            if (name.Length < 1)
            {
                return;
            }

            if (!checkedListBox_Filters.Items.Contains(name))
            {
                SortedSet<string> names = new SortedSet<string>();
                names.Add(name);
                foreach (var objName in checkedListBox_Filters.Items)
                {
                    names.Add((string)objName);
                }

                HashSet<string> checkedNames = new HashSet<string>();
                checkedNames.Add(name);
                foreach (var objName in checkedListBox_Filters.CheckedItems)
                {
                    checkedNames.Add((string)objName);
                }

                filterNames.Clear();
                checkedListBox_Filters.Items.Clear();
                foreach (var tempName in names)
                {
                    int itemIndex = checkedListBox_Filters.Items.Add(tempName);
                    if (checkedNames.Contains(tempName))
                    {
                        checkedListBox_Filters.SetItemChecked(itemIndex, true);
                    }
                }
            }
        }

        private void SetLanguage(string language)
        {
            if (language == SecretLanguage.German)
            {
                SecretLanguage.SetGerman();
            }
            else if (language == SecretLanguage.French)
            {
                SecretLanguage.SetFrench();
            }
            else
            {
                SecretLanguage.SetEnglish();
            }
        }

        #region UI Events
        private void checkBox_Filter_CheckedChanged(object sender, EventArgs e)
        {
            SetFilterEnabled(checkBox_Filter.Checked);
        }

        private void textBox_FilterName_Enter(object sender, EventArgs e)
        {
            AddFilterName(textBox_FilterName.Text);
        }

        private void textBox_FilterName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                AddFilterName(textBox_FilterName.Text);
                textBox_FilterName.Text = "";
            }
        }

        private void button_AddCharacter_Click(object sender, EventArgs e)
        {
            AddFilterName(textBox_FilterName.Text);
            textBox_FilterName.Text = "";
        }

        private void button_DeleteCharacter_Click(object sender, EventArgs e)
        {
            int index = checkedListBox_Filters.SelectedIndex;
            if (index >= 0)
            {
                var objItem = checkedListBox_Filters.SelectedItem;
                string strItem = (string)objItem;
                filterNames.Remove(strItem);
                checkedListBox_Filters.Items.Remove(objItem);

                if (index < checkedListBox_Filters.Items.Count)
                {
                    checkedListBox_Filters.SelectedIndex = index;
                }
            }
        }

        private void checkedListBox_Filters_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (checkBox_Filter.Checked && e.CurrentValue != e.NewValue)
            {
                if (e.NewValue == CheckState.Checked)
                {
                    filterNames.Add((string)checkedListBox_Filters.Items[e.Index]);
                }
                else
                {
                    filterNames.Remove((string)checkedListBox_Filters.Items[e.Index]);
                }
            }
        }

        private void checkBox_EnableTSWAddon_CheckedChanged(object sender, EventArgs e)
        {
            lock (addonEnabledLocker)
            {
                addonEnabled = checkBox_EnableTSWAddon.Checked;
            }
        }

        private void checkBox_ExportScript_CheckedChanged(object sender, EventArgs e)
        {
            SetExportFieldStatus();
        }

        private void comboBox_Language_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetLanguage(comboBox_Language.Text);
        }
        #endregion

        #region Help Texts
        private void checkBox_AutoCheck_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Check for updates\n\n" +
                                   "The plugin will every time it loads check the web page for new updates. " +
                                   "If new update is found, then\n" +
                                   "it will notice the user about the new version. Note that it will not update the plugin " +
                                   "automaticly.");
        }

        private void label1_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Information about the Secret Combat Parser Plugin.\n\n" +
                                   "Remember to do a /logcombat on inside The Secret World game client to start the logging!");
        }

        private void checkBox_ExportScript_MouseHover(object sender, EventArgs e)
        {
            string helpPrefix = "Export summary to /actchat script\n\n" +
                                   "Secret World allows users to run scripts to perform simple actions. " +
                                   "This option causes ACT to generate a script with the latest parse details.\n" +
                                   "Just type /actchat in game and your summary will appear.\n";

            if (checkBox_ExportScript.Checked && !Directory.Exists(GetScriptFolder()))
            {
                ActGlobals.oFormActMain.SetOptionsHelpText(helpPrefix + "\nCREATE A \"Scripts\" FOLDER UNDER YOUR SECRET WORLD GAME FOLDER FOR THIS OPTION TO WORK");
            }
            else
            {
                ActGlobals.oFormActMain.SetOptionsHelpText(helpPrefix);
            }
        }

        private void checkBox_LimitNames_MouseHover(object sender, EventArgs e) {
            ActGlobals.oFormActMain.SetOptionsHelpText("Limit Player-Names\n\n" +
                                   "Limits the playernames to 7 characters using /actchat.");
        }

        private void checkBox_ExportRoundDPS_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Round DPS / HPS values (xxx instead of xxx.xx).");
        }

        private void checkBox_ExportShowLegend_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Shows the legend in chat exports.");
        }

        private void checkBox_ExportColored_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Coloring the chat export.");
        }

        private void checkBox_ExportSplit_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Splits the chat export (not the acttell!) into seperate scripts if necesarry.");
        }

        private void checkBox_ExportHtml_MouseHover(object sender, EventArgs e)
        {
            string helpPrefix = "Export HTML summary to /act script\n\n" +
                                   "Secret World allows users to run scripts to perform simple actions. " +
                                   "This option causes ACT to generate a script with the latest parse details.\n" +
                                   "Just type /act in game and your summary will appear.\n";

            if (checkBox_ExportHtml.Checked && !Directory.Exists(GetScriptFolder()))
            {
                ActGlobals.oFormActMain.SetOptionsHelpText(helpPrefix + "\nCREATE A \"Scripts\" FOLDER UNDER YOUR SECRET WORLD GAME FOLDER FOR THIS OPTION TO WORK");
            }
            else
            {
                ActGlobals.oFormActMain.SetOptionsHelpText(helpPrefix);
            }
        }

        private void checkBox_ExportAllies_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Export you team/raid members to the /act and /actchat scripts\n\n" +
                                    "This option will cause the output scripts to filter out any values for non-player characters.\n" +
                                    "Uncheck it to show all mob values in your scripts.");
        }

        private void checkBox_DontExportShortEnc_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Dont export short encounters to the /act and /actchat scripts\n\n" +
                                    "This option will cause the output scripts to filter out any encounters shorter than 10 seconds.");
        }

        private void checkBox_EnableTSWAddon_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Enable TSWACT Addon features\n\n" +
                                   "The Secret World combat log does not include messages about group members or combat start/stop.  " +
                                   "The TSWACT in-game addon logs these addition messages into the ClientLog.txt file.\n\n" +
                                   "You can download the addon from SecretUI or Curse.");
        }

        private void checkBox_AutofixDBConf_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Autofix the required addon settings in the TSW dbDebug.conf file.\n\n" +
                                    "The TSWACT addon requires two values to be set in the game's dbDebug.conf file.  Unfortunately the patcher keeps reseting these values.\n" +
                                    "This option will get the ACT plugin to auto-fix the file, but it will only work if you start ACT before starting TSW.");
        }

        private void checkBox_HideUnknown_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Hide TSW_Unknown entries\n\n" +
                                   "Secret World sometimes writes entries in the combat log without the name of the damage dealer.  " +
                                   "These entries are shown as TSW_Unknown in ACT.  Select this option to hide these entries.");
        }

        private void checkBox_SelfDamage_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Show self-damage\n\n" +
                                   "Selecting this option generates a new character '_self_', which deals the correct damage to the chars.\n" +
                                   "This will affect hit-count based stats like crit% and pen%.\n\n" +
                                   "When unselected, the damage done to yourself is set to 0.");
        }

        private void checkBox_SelfPlayerDamage_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Show by player\n\n" +
                                   "If checked, the generated character is '<playername>_self'.");
        }

        private void checkBox_ReduceAegis_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Reduce Aegis-Names\n\n" +
                                   "If checked, the name of the victim with active aegis shields will be reduced to it's name\n" +
                                   "without aegis. The damagetype is set to the corresponding aegis-shield name.\n\n" +
                                   "Otherwise the victim might appear several times, one for each aegis type.\n" +
                                   "The damagetype is set to Aegis when hitting aegis-shields.");
        }

        private void checkBox_Filter_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Enable/disable character name filtering\n\n" +
                                   "Only the results for the checked character names will show up in ACT when this is enabled.");
        }

        private void checkedListBox_Filters_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("List of character names to filter in ACT\n\n" +
                                   "All the names that are checked in this list will be shown in the ACT results.\n" +
                                   "Result information for all other characters and NPCs will be discarded.");
        }

        private void textBox_FilterName_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("The name of a new character to add to the filter list\n\n" +
                                   "Enter the name of the character you want to include and hit the Enter key or the Add button.");
        }

        private void button_AddCharacter_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Adds the character name from the above text box to the filter list.");
        }

        private void button_DeleteCharacter_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Deletes the currently selected character from the filter list.");
        }

        private void checkBox_filterExclude_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Causes the filtering to hide entries from the selected characters, rather than show their entries.");
        }

        private void checkBox_filterScript_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Causes the filtering to filter the entries in the generated /act script.");
        }

        private void comboBox_Language_MouseHover(object sender, EventArgs e)
        {
            ActGlobals.oFormActMain.SetOptionsHelpText("Specifies the language of the Secret World client.");
        }
        #endregion

        private void SetExportFieldStatus()
        {
            checkedListBox_ExportFields.Enabled = checkBox_ExportScript.Checked;
        }

    }

    public class BuffData
    {
        public DateTime StartTime { get; private set; }
        public double TotalTime { get; private set; }
        public bool Changed { get; private set; }

        public BuffData()
        {
            StartTime = DateTime.MaxValue;
            TotalTime = 0;
            Changed = false;
        }

        public void Start(DateTime currentTime)
        {
            if (StartTime == DateTime.MaxValue)
            {
                StartTime = currentTime;
            }

            Changed = true;
        }

        public void Restart(DateTime currentTime)
        {
            StartTime = currentTime;
            TotalTime = 0;
            Changed = false;
        }

        public void Stop(DateTime currentTime)
        {
            if (StartTime != DateTime.MaxValue)
            {
                TimeSpan diff = currentTime.Subtract(StartTime);
                if (diff.TotalMilliseconds > 0)
                {
                    TotalTime = TotalTime + (long)diff.TotalMilliseconds;
                }

                StartTime = DateTime.MaxValue;
                Changed = true;
            }
        }

        public bool IsRunning()
        {
            return !(StartTime == DateTime.MaxValue);
        }

        public BuffData Copy()
        {
            BuffData ret = new BuffData();
            ret.StartTime = this.StartTime;
            ret.TotalTime = this.TotalTime;
            ret.Changed = this.Changed;
            return ret;
        }
    }

    // Altuslumen 1.4.0.0 start
    public static class SecretLanguage
    {
        // Supported languages
        public static string English = "English";
        public static string German = "German";
        public static string French = "French";

        // Current language
        public static string Language = English;

        // Special damage types
        public static string none = "";
        public static string Unknown = "";

        // Additional parsing strings
        public static string Killing = "";
        public static string Miss = "";
        public static HashSet<string> YouSet;
        public static string You = "";
        public static string Physical = "";
        public static string Glancing = "";
        public static string Penetrated = "";
        public static string Blocked = "";
        public static string Aegis = "";

        // Filter string
        public static string HitLine1 = "";
        public static string HitLine2 = "";
        public static string HateLine = "";
        public static string HealLine1 = "";
        public static string HealLine2 = "";
        public static string AbsorbLine = "";
        public static string EvadeLine1 = "";
        public static string EvadeLine2 = "";
        public static string DiedLine = "";
        public static string RedirectedLine = "";
        public static string BuffLine1 = "";
        public static string BuffLine2 = "";
        public static string BuffStopLine = "";
        public static string AegisShieldLine = "";

        // Regex
        public static List<Regex> damageLines;
        public static List<Regex> redirectLines;
        public static List<Regex> healLines;
        public static List<Regex> evadedLines;
        public static List<Regex> diedLines;
        public static List<Regex> buffLines;
        public static List<Regex> buffStopLines;
        public static List<Regex> aegisShieldLines;

        // Whisper command
        public static string WhisperCmd = "";

        // Correcting the origin of those skills
//        public static HashSet<string> DamageWithoutOrigin;

        public static void SetEnglish()
        {
            Physical = "physical";
            Glancing = "Glancing";
            Penetrated = "Penetrated";
            Blocked = "Blocked";
            Aegis = "Aegis";
            RedirectedLine = " redirects ";
            HitLine1 = " hits ";
            HitLine2 = " dealt ";
            HateLine = " successfully used ";
            HealLine1 = " heals ";
            HealLine2 = " restored ";
            AbsorbLine = " absorbs ";
            EvadeLine1 = " evade";
            EvadeLine2 = "NotUsedHere";
            DiedLine = " died.";
            BuffLine1 = " buff ";
            BuffLine2 = " affects ";
            BuffStopLine = "Buff";
            AegisShieldLine = "AEGIS";
            WhisperCmd = "/w";
            You = "you";
            YouSet = new HashSet<string>();
            YouSet.Add(You);
            YouSet.Add("your");

//            DamageWithoutOrigin = new HashSet<string>();

            damageLines = new List<Regex>();
            string apostropheSkills = "Thor's Hammer|Nightmare's Edge|Miner's Claw|Gaia's Presence|Carter's Burning|Angel's Touch|Adam's Rib|Dream's End|Stone's Throw|Mjolnir's Echo|Shadow's Shadow|The Carver's Art|Dragon's Breath|The Inspector's Gadget|Protector's Wrath|Keziah Mason's Ring|Extradimensional Doppelganger's Dissonance|Extradimensional Doppelganger's Dissipation|Extradimensional Doppelganger's Cacophony|Beginner's Luck|Snake's Bite|Gambler's Soul|Warrior's Spirit|Skadi's Ring";
            damageLines.Add(new Regex(@"^(?<actor>Your|.*'s)\s(?<attackName>" + apostropheSkills + @")\sdealt\s(?<amount>[0-9]+)\sdamage\sto\s(?<actee>.+)\.$", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<actor>Your|.*'s)\s(?<attackName>.+?)\sdealt\s(?<amount>[0-9]+)\sdamage\sto\s(?<actee>.+)\.$", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<actor>Your|.*'s)\s(?<attackName>" + apostropheSkills + @")\shits\s(?<damageType>\([^\)]+\)\s)?(?<actee>.+)\sfor\s(?<amount>[0-9]+)(?<damageClass>.*?)(?<ignore>\sdamage)?\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<actor>Your|.*'s)\s(?<attackName>.+)\shits\s(?<damageType>\([^\)]+\)\s)?(?<actee>.+)\sfor\s(?<amount>[0-9]+)(?<damageClass>.*?)(?<ignore>\sdamage)?\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<attackName>" + apostropheSkills + @")\shits\s(?<damageType>\([^\)]+\)\s)?(?<actee>.+)\sfor\s(?<amount>[0-9]+)(?<damageClass>.*?)(?<ignore>\sdamage)?\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<attackName>.+)\shits\s(?<damageType>\([^\)]+\)\s)?(?<actee>.+)\sfor\s(?<amount>[0-9]+)(?<damageClass>.*?)(?<ignore>\sdamage)?\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<actor>Your)\s(?<attackName>.+?)\s(?<crit>critically\s)hits\s(?<actee>.+)\sfor\s(?<amount>[0-9]+)\.", RegexOptions.Compiled));
            string hateSkills = "Provoke|Mass Provocation|Confuse|Mass Confusion|Misdirection";
            damageLines.Add(new Regex(@"^(?<actor>.+)\ssuccessfully\sused\sthe\s(?<attackName>" + hateSkills + @")\.", RegexOptions.Compiled));


            redirectLines = new List<Regex>();
            redirectLines.Add(new Regex(@"^(?<actor>Your|.*'s)\s(?<attackName>" + apostropheSkills + @")\sredirects\s(?<amount>[0-9]+)\sdamage\sto\s(?<actee>.*)\.$", RegexOptions.Compiled));
            redirectLines.Add(new Regex(@"^(?<actor>Your|.*'s)\s(?<attackName>.+)\sredirects\s(?<amount>[0-9]+)\sdamage\sto\s(?<actee>.*)\.$", RegexOptions.Compiled));

            healLines = new List<Regex>();
            healLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<actor>Your|.*'s)\s(?<attackName>" + apostropheSkills + @")\sheals\s(?<actee>.+)\sfor\s(?<amount>[0-9]+)\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<actor>Your|.*'s)\s(?<attackName>.+)\sheals\s(?<actee>.+)\sfor\s(?<amount>[0-9]+)\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<actor>Your|.*'s)\s(?<attackName>.+)\srestored\s(?<actee>.+)\sby\s(?<amount>[0-9]+)\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<actor>Your|.*'s)\s(?<attackName>.+)\sabsorbs\s(?<amount>[0-9]+)\sdamage\sfrom\s(?<actee>.+)'s.*\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critical\)\s)?(?<actor>Your|.*'s)\s(?<attackName>" + apostropheSkills + @")\sabsorbs\s(?<amount>[0-9]+)\sdamage\sfrom\s(?<actee>.+)'s.*\.", RegexOptions.Compiled));

            evadedLines = new List<Regex>();
            evadedLines.Add(new Regex(@"^(?<actee>.+)\sevaded\s(?<actor>your|.*'s)\s(?<attackName>" + apostropheSkills + @")\.", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actee>.+)\sevaded\s(?<actor>your|.*'s)\s(?<attackName>.+)\.", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actee>You)\sevade\s(?<actor>.*'s)\s(?<attackName>" + apostropheSkills + @")\.", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actee>You)\sevade\s(?<actor>.*'s)\s(?<attackName>.+)\.$", RegexOptions.Compiled));

            diedLines = new List<Regex>();
            diedLines.Add(new Regex(@"^(?<actor>.*)\sdied\.", RegexOptions.Compiled));

            buffLines = new List<Regex>();
            buffLines.Add(new Regex(@"^You\sgain\sbuff\s(?<name>.*)$", RegexOptions.Compiled));
            buffLines.Add(new Regex(@".*\saffects\syou\swith\s(?<name>.*)$", RegexOptions.Compiled));

            buffStopLines = new List<Regex>();
            buffStopLines.Add(new Regex(@"^Buff\s(?<name>.*)\sterminated\.", RegexOptions.Compiled));

            aegisShieldLines = new List<Regex>();
            aegisShieldLines.Add(new Regex(@"^(?<actee>.+)'s\s(?<aegis>.+)\sAEGIS", RegexOptions.Compiled));
            aegisShieldLines.Add(new Regex(@"^(?<actee>your)\s(?<aegis>.+)\sAEGIS", RegexOptions.Compiled));

            // Current language
            Language = English;

            // Special damage types
            none = "none";
            Unknown = "Unknown";

            // Additional parsing strings
            Killing = "Death";
            Miss = "Miss";
        }


        public static void SetGerman()
        {
            // Sets the Language variable to be referenced.
            Language = German;

            // Sets specific German values to override the defaults
            Physical = "Körperlich";
            Glancing = "Streiftreffer";
            Penetrated = "Durchdringend";
            Blocked = "Geblockt";
            Aegis = "Aegis";
            RedirectedLine = " leitet ";
            HitLine1 = "Schaden";
            HitLine2 = "NotUsedHere";
            HateLine = " erfolgreich ";
            HealLine1 = " Heilung.";
            HealLine2 = " regeneriert ";
            AbsorbLine = " absorbiert ";
            EvadeLine1 = " ausgewichen";
            EvadeLine2 = "NotUsedHere";
            DiedLine = " ist tot.";
            BuffLine1 = " Buff ";
            BuffLine2 = " belegt ";
            BuffStopLine = "Buff ";
            AegisShieldLine = "AEGIS";
            WhisperCmd = "/w";
            You = "ihnen";
            YouSet = new HashSet<string>();
            YouSet.Add(You);
            YouSet.Add("ihr");
            YouSet.Add("sie");
            YouSet.Add("ihre");
            YouSet.Add("ihrer");
            YouSet.Add("ihrem");
            YouSet.Add("ihres");

//            DamageWithoutOrigin = new HashSet<string>();
//            DamageWithoutOrigin.Add("Schmutz");
//            DamageWithoutOrigin.Add("Elektrifiziert");
//            DamageWithoutOrigin.Add("Brennt");
//            DamageWithoutOrigin.Add("Kochend heiß");
//            DamageWithoutOrigin.Add("In Flammen");
//            DamageWithoutOrigin.Add("Superheißes Metall");
//            DamageWithoutOrigin.Add("Lebensverbrennung");
//            DamageWithoutOrigin.Add("Traumgewand");
//            DamageWithoutOrigin.Add("Moloch");
//            DamageWithoutOrigin.Add("Schmutzsog");
//            DamageWithoutOrigin.Add("Entfernen");
//            //DamageWithoutOrigin.Add("Auflösung"); //damage type: filth // Can't do this here because of the blood skill with the same name
//            DamageWithoutOrigin.Add("Herabstürzende Trümmer");
//            DamageWithoutOrigin.Add("Konditionierung");
//            DamageWithoutOrigin.Add("Hauruck!");
//            DamageWithoutOrigin.Add("Angriff");
            damageLines = new List<Regex>();
            string apostropheSkills = "Tod von oben|Androhung von Waffengewalt|Von Anfang bis Ende|Runter von meinem Land|Sturm von Niflheim|Unterstützung von außen";
	//SWL
			//(Kritischer Treffer)\"Anfängerglück\" (Normal) trifft KrustenBraten und verursacht 1023 Körperlich-Schaden.
			//(Kritischer Treffer) \"Streuschuss\" (Normal) trifft Martling und verursacht 1834 Körperlich-Schaden.
            damageLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s|\(Kritischer Treffer\)?)?(?<attackName>.+)\s(?<damageType>\([^\)]+\))\strifft\s(?<actee>.+)\sund\sverursacht\s(?<amount>[0-9]+)\s(?<damageClass>.*?)-Schaden\."));
            //"Vollautomatik" (Normal) trifft Untoter Inselbewohner und verursacht 55 Körperlich-Schaden
            damageLines.Add(new Regex(@"^(?<attackName>.+)\s(?<damageType>\([^\)]+\))\strifft\s(?<actee>.+)\sund\sverursacht\s(?<amount>[0-9]+)\s(?<damageClass>.*?)-Schaden", RegexOptions.Compiled));
            //TheOtherPlayer trifft mit "Kontrolliertes Feuern" (Normal) und verursacht Untoter Inselbewohner 250 Körperlich-Schaden.
            damageLines.Add(new Regex(@"^(?<actor>.+)\strifft\smit\s(?<attackName>.+)\s(?<damageType>\([^\)]+\))\sund\sverursacht\s(?<actee>.+)\s(?<amount>[0-9]+)\s(?<damageClass>.*?)-Schaden\.", RegexOptions.Compiled));
            //(Kritischer Treffer) TheOtherPlayer trifft mit \"Kontrolliertes Feuern\" (Normal) Untoter Inselbewohner und verursacht 336 Körperlich-Schaden.
            damageLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s|\(Kritischer Treffer\)\s)?(?<actor>.+)\strifft\smit\s(?<attackName>.+)\s(?<damageType>\([^\)]+\))\s(?<actee>.+)\sund\sverursacht\s(?<amount>[0-9]+)\s(?<damageClass>.*?)-Schaden\.", RegexOptions.Compiled));
			 //(Kritischer Treffer) TheOtherPlayer trifft Sie mit \"Flackern\" (Normal) und verursacht 929 Magie-Schaden.
            damageLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s|\(Kritischer Treffer\)\s)?(?<actor>.+)\strifft\s(?<actee>.+?)mit\s(?<attackName>.+)\s(?<damageType>\([^\)]+\))\sund\sverursacht\s(?<amount>[0-9]+)\s(?<damageClass>.*?)-Schaden\.", RegexOptions.Compiled));
            //Haugbui-Jarl trifft Sie mit "Aufgeladener Hack" (Normal) und verursacht 84 Magie-Schaden.
            damageLines.Add(new Regex(@"^(?<actor>.+) trifft (?<actee>.+?) mit (?<attackName>.+?) (?<damageType>\([^\)]+\)) und verursacht (?<amount>[0-9]+) (?<damageClass>.*?)-Schaden\.", RegexOptions.Compiled));
            //Elektrifiziert trifft Sie (Normal) und verursacht 212 Magie-Schaden.
            damageLines.Add(new Regex(@"(?<actor>.+) trifft (?<actee>.+?) (?<damageType>\([^\)]+\)) und verursacht (?<amount>[0-9]+) (?<damageClass>.*?)-Schaden\.", RegexOptions.Compiled));
			
	//TSW
            damageLines.Add(new Regex(@"^(?<actor>Ihre)\s(?<attackName>.+?)-Kraft\sfügt\s(?<actee>.+)\s(?<amount>[0-9]+)\sSchaden\szu\.", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s|\(Kritischer Treffer\)\s)?(?<attackName>(""?)(" + apostropheSkills + @")(""?))\s(?<damageType>\([^\)]+\)\s)?fügt\s(?<actee>.+)\s(?<amount>[0-9]+)(?<damageClass>.*?)(\-)?(?<ignore>Schaden\szu)?\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s|\(Kritischer Treffer\)\s)?(?<attackName>(""?)(" + apostropheSkills + @")(""?))\s(?<damageType>\([^\)]+\)\s)?von\s(?<actor>.+?)\sfügt\s(?<actee>.+?)\s(?<amount>[0-9]+)(?<damageClass>.*?)(\-)?(?<ignore>Schaden\szu)?\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s|\(Kritischer Treffer\)\s)?(?<attackName>.+?)\s(?<damageType>\([^\)]+\)\s)?von\s(?<actor>.+?)\sfügt\s(?<actee>.+)\s(?<amount>[0-9]+)(?<damageClass>.*?)(\-)?(?<ignore>Schaden\szu)?\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s|\(Kritischer Treffer\)\s)?(?<attackName>.+?)\s(?<damageType>\([^\)]+\)\s)?fügt\s(?<actee>.+)\s(?<amount>[0-9]+)(?<damageClass>.*?)(?<ignore>-Schaden\szu)?\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<actor>Ihr)\s(?<attackName>.+?)\strifft\s(?<actee>.+?)\s(?<crit>kritisch\s)?für\s(?<amount>[0-9]+)\sSchaden\.", RegexOptions.Compiled));
	//TSW + SWL
			string hateSkills = "Provozieren|Massenprovokation|Verwirren|Massenverwirrung|Irreführung";
            damageLines.Add(new Regex(@"^(?<actor>.+)\s(haben|hat)\serfolgreich\s'(?<attackName>" + hateSkills + @")'\seingesetzt\.", RegexOptions.Compiled));

            redirectLines = new List<Regex>();
			//Blast Shield von Feuer von Sheol leitet 20 Schaden auf Tipsu um.
            redirectLines.Add(new Regex(@"^(?<attackName>" + apostropheSkills + @"?)\s(?<damageType>\([^\)]+\)\s)?von\s(?<actor>.+?)\sleitet\s(?<amount>[0-9]+)\sSchaden\sauf\s(?<actee>.+?)\sum\.", RegexOptions.Compiled));
            redirectLines.Add(new Regex(@"^(?<attackName>.+?)\s(?<damageType>\([^\)]+\)\s)?von\s(?<actor>.+?)\sleitet\s(?<amount>[0-9]+)\sSchaden\sauf\s(?<actee>.+?)\sum\.", RegexOptions.Compiled));
            redirectLines.Add(new Regex(@"^(?<actor>Ihr)\s(?<attackName>.+?)\s(?<damageType>\([^\)]+\)\s)?leitet\s(?<amount>[0-9]+)\sSchaden\sauf\s(?<actee>.+?)\sum\.", RegexOptions.Compiled));
            redirectLines.Add(new Regex(@"^(?<actor>Ihr)\s(?<attackName>" + apostropheSkills + @"?)\s(?<damageType>\([^\)]+\)\s)?leitet\s(?<amount>[0-9]+)\sSchaden\sauf\s(?<actee>.+?)\sum\.", RegexOptions.Compiled));

            healLines = new List<Regex>();
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?("")?(?<attackName>" + apostropheSkills + @")("")?\svon\s(?<actor>.*?)\sgewährt\s(?<actee>.+?)\s(?<amount>[0-9]+)\sPunkte\sHeilung\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?(?<attackName>.+?)\svon\s(?<actor>.+?)?(\s)?gewährt\s(?<actee>.+?)\s(?<amount>[0-9]+)\sPunkte\sHeilung\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?(Ihr\s)?(?<attackName>.+?)\sgewährt\s(?<actee>.+?)\s(?<amount>[0-9]+)\sPunkte\sHeilung\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?(?<attackName>.+?)\sgewährt\s(?<actee>.+?)\s(?<amount>[0-9]+)\sHeilung\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?(?<attackName>.+?)(-Kraft)?\svon\s(?<actor>.*?)\sregeneriert\s(?<amount>[0-9]+)\s(der\s)?(?<actee>.+?)\.$", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?(?<actor>Ihre)\s(?<attackName>.+?)\sregeneriert\s(?<amount>[0-9]+)\s(der\s)?(?<actee>.+?)\.$", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?(?<actor>.+?)\sabsorbiert\s(?<amount>[0-9]+)\sSchaden\s(?<actee>Ihres)\s(?<attackName>" + apostropheSkills + @"?)\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?(?<actor>.+?)\sabsorbiert\s(?<amount>[0-9]+)\sSchaden\s(?<actee>Ihres)\s(?<attackName>.+?)\.$", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Kritisch\)\s)?(?<actor>Ihr)\s(?<attackName>.+?)\sabsorbiert\s(?<amount>[0-9]+)\sSchaden\sdes\sAngriffs\svon\s(?<actee>.+?)\.$", RegexOptions.Compiled));

            evadedLines = new List<Regex>();
            evadedLines.Add(new Regex(@"^(?<actee>.+?)\sist\s(?<actor>Ihrer)\sKraft\s(?<attackName>.+?)\sausgewichen\.$", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actee>.+?)\sist\s("")?(?<attackName>" + apostropheSkills + @")("")?\svon\s(?<actor>.+?)\sausgewichen\.$", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actor>Sie)\ssind\s("")?(?<attackName>" + apostropheSkills + @")("")?\svon\s(?<actor>.+?)\sausgewichen\.$", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actee>.+?)\sist\s(?<attackName>.+?)\svon\s(?<actor>.+?)\sausgewichen\.$", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actor>Sie)\ssind\s(?<attackName>.+?)\svon\s(?<actor>.+?)\sausgewichen\.$", RegexOptions.Compiled));

            diedLines = new List<Regex>();
            diedLines.Add(new Regex(@"^(?<actor>.*)\sist\stot\.", RegexOptions.Compiled));

            buffLines = new List<Regex>();
            buffLines.Add(new Regex(@"^Sie\serhalten\sden\sBuff\s(?<name>.*)\.$", RegexOptions.Compiled));
            buffLines.Add(new Regex(@"\sbelegt\sSie\smit\s(?<name>.*)$", RegexOptions.Compiled));

            buffStopLines = new List<Regex>();
            buffStopLines.Add(new Regex(@"^Buff\s(?<name>.*)\sbeendet\.", RegexOptions.Compiled));

            aegisShieldLines = new List<Regex>();
            aegisShieldLines.Add(new Regex(@"(?<aegis>.+)\-AEGIS\svon\s(?<actee>.+)", RegexOptions.Compiled));
            aegisShieldLines.Add(new Regex(@"(?<actee>Ihrem|Ihres)\s(?<aegis>.+)\-AEGIS", RegexOptions.Compiled));

            // Special damage types
            none = "none";
            Unknown = "Unknown";

            // Additional parsing strings
            Killing = "Death";
            Miss = "Miss";
        }

        public static void SetFrench()
        {
            // Sets the Language variable to be referenced.
            Language = French;

            // Sets specific French values to override the defaults
            Physical = "physiques";
            Glancing = "Coup dévié";
            Penetrated = "Passé";
            Blocked = "Paré";
            Aegis = "Aegis";
            RedirectedLine = " redirige ";
            HitLine1 = " inflige";
            //HitLine2 = "NotUsedHere";
            //TSW Language Mixup Hack
            HitLine2 = " dealt ";
            HateLine = "succès";
            HealLine1 = " rend ";
            HealLine2 = " résistance ";
            AbsorbLine = " NOT-KNOWN ";
            EvadeLine1 = " évité ";
            EvadeLine2 = " évite ";
            DiedLine = " a perdu la vie.";
            BuffLine1 = " buff : ";
            BuffLine2 = " affecte ";
            BuffStopLine = "Buff";
            AegisShieldLine = "AEGIS";
            WhisperCmd = "/dire";
            You = "vous";
            YouSet = new HashSet<string>();
            YouSet.Add(You);
            YouSet.Add("votre");
            //TSW Language Mixup Hack
            YouSet.Add("your");

            //            DamageWithoutOrigin = new HashSet<string>();

            // The regex can't differentiate between attacks with "de" in the name and characters with "de" in their name :(
            SortedSet<string> deAttacks = new SortedSet<string>();
            SortedSet<string> deAttacks2 = new SortedSet<string>();
            SortedSet<string> deAttacks3 = new SortedSet<string>();
         //TSW
			// main-actifs
            deAttacks.Add("Avènement de la discorde");
            deAttacks.Add("Ballet de balles");
            deAttacks.Add("Bouclier de sang");
            deAttacks.Add("Boule de feu");
            deAttacks.Add("Brûlure de poudre");
            deAttacks.Add("Chance de débutant");
            deAttacks.Add("Charge de la brousse");
            deAttacks.Add("Cinq pétales de lotus");
            deAttacks.Add("Coup de boutoir");
            deAttacks.Add("Coup de fouet");
            deAttacks.Add("Coup de pompe");
            deAttacks.Add("Coup de sang");
            deAttacks.Add("Cran de sûreté ôté");
            deAttacks.Add("Croc du croissant de lune");
            deAttacks.Add("Déferlement de lames");
            deAttacks.Add("Drone de garde");
            deAttacks.Add("Drone de renfort");
            deAttacks.Add("Drone de secteur");
            deAttacks.Add("Écho de l'acier");
            deAttacks.Add("Éclats de rasoir");
            deAttacks.Add("État de choc");
            deAttacks.Add("Enfer de feu");
            deAttacks.Add("Envie de tuer");
            deAttacks.Add("Excité de la gâchette");
            deAttacks.Add("Fracture de la réalité");
            deAttacks.Add("Grand coup de marteau");
            deAttacks.Add("Inversion de tendance");
            deAttacks.Add("Jet de pierre");
            deAttacks.Add("L'art de la guerre");
            deAttacks.Add("Manifestation de feu");
            deAttacks.Add("Manifestation de foudre");
            deAttacks.Add("Manifestation de glace");
            deAttacks.Add("Marteau de Thor");
            deAttacks.Add("Moisson de sang");
            deAttacks.Add("Onde de choc");
            deAttacks.Add("Œil de la tempête");
            deAttacks.Add("Poussée de sang");
            deAttacks.Add("Rafale de trois balles");
            deAttacks.Add("Ralentissement de l'ennemi");
            deAttacks.Add("Rayon de glace");
            deAttacks.Add("Rayon de ponction");
            deAttacks.Add("Règlement de comptes");
            deAttacks.Add("Récupération de cartouches");
            deAttacks.Add("Rideau de rubis");
            deAttacks.Add("Roue de poignards");
            deAttacks.Add("Salve de cartouches");
            deAttacks.Add("Sang de guerre");
            deAttacks.Add("Symbole de terreur");
            deAttacks.Add("Taille de bambou");
            deAttacks.Add("Taille de brindille");
            deAttacks.Add("Tête de pont");
            deAttacks.Add("Tir de couverture");
            deAttacks.Add("Tir de précision");
            deAttacks.Add("Trait de feu");
            deAttacks.Add("Une de chaque");
            deAttacks.Add("Union de décharges");
            deAttacks.Add("Vague de froid");
            deAttacks.Add("Vague de gel");
            // main-passifs
            deAttacks.Add("Coup de grâce");
            deAttacks.Add("Du fond de l'abîme");
            deAttacks.Add("Effet de souffle");
            deAttacks.Add("Effet de terre");
            deAttacks.Add("Mépris de la douleur");
            deAttacks.Add("Pouvoir de l'esprit");
            deAttacks.Add("Tactique de guérilla");
            deAttacks.Add("Tête de pont renforcée");
            deAttacks.Add("Volée de dagues");
            // aux-actifs
            deAttacks.Add("L'art de la découpe");
            deAttacks.Add("Déluge de coups de fouet");
            deAttacks.Add("Odeur de victoire-<font color='#DDDDDD'>Du feu de dieu</font>");
            deAttacks.Add("<font color='#DDDDDD'>Odeur de victoire</font>-Du feu de dieu");
            // aux-passifs
            deAttacks.Add("Claquement de fouet");
            // ennemies
            deAttacks.Add("Absorption de Gaia");
            deAttacks.Add("Appel de la vermine");
            deAttacks.Add("Atteinte de Méphisto");
            deAttacks.Add("Balles de feu grégeois");
            deAttacks.Add("Braillement de mort");
            deAttacks.Add("Champ de transfusion");
            deAttacks.Add("Concentration de tueur");
            deAttacks.Add("Coup de laisse");
            deAttacks.Add("Coupure de courant");
            deAttacks.Add("Courant de Souillure");
            deAttacks.Add("Décharge de fusil à pompe");
            deAttacks.Add("Dégagement de salle");
            deAttacks.Add("Dégâts de dégel");
            deAttacks.Add("Du sang pour le chien de sang");
            deAttacks.Add("Écho de Mjolnir");
            deAttacks.Add("Embrasement de forge");
            deAttacks.Add("Éruption de lave");
            deAttacks.Add("Éruptions de fièvre");
            deAttacks.Add("Explosion de conteneur cryogénique");
            deAttacks.Add("Explosion de viande");
            deAttacks.Add("Fente de la légion");
            deAttacks.Add("Feux de Phlegethos");
            deAttacks.Add("Fin de rêve");
            deAttacks.Add("Force de l'au-delà");
            deAttacks.Add("Geyser de sable");
            deAttacks.Add("Gravure de Cucuvea");
            deAttacks.Add("Houle de mutation");
            deAttacks.Add("Hors de mes terres");
            deAttacks.Add("Invocation de corbeaux");
            deAttacks.Add("Invocation de l'immolation");
            deAttacks.Add("Lamelle de vie");
            deAttacks.Add("Lance de mépris");
            deAttacks.Add("Laser de pistage");
            deAttacks.Add("Libération de conteneur");
            deAttacks.Add("Messe de minuit");
            deAttacks.Add("Mine de proximité explosive");
            deAttacks.Add("Nuée de sangsues");
            deAttacks.Add("Offrande de sang");
            deAttacks.Add("Ouverture de faille");
            deAttacks.Add("Petit sort d'aura de feu");
            deAttacks.Add("Pluie de balles");
            deAttacks.Add("Prisonnier de la glace");
            deAttacks.Add("Psaume de démêlage");
            deAttacks.Add("Pulsation de phase");
            deAttacks.Add("Pulvérisation de prière");
            deAttacks.Add("Refroidisseur de pulsations");
            deAttacks.Add("Refus de zone");
            deAttacks.Add("Rempart de fer noir");
            deAttacks.Add("Résidu de vase toxique");
            deAttacks.Add("Roue de douleur surmultipliée");
            deAttacks.Add("Route de Xibalba");
            deAttacks.Add("Ruade de cerf");
            deAttacks.Add("Sacrifice de sang des ak'abs");
            deAttacks.Add("Souffle de fournaise");
            deAttacks.Add("Tempête de métal");
            deAttacks.Add("Tempête de sable");
            deAttacks.Add("Tir de glace");
            deAttacks.Add("Tir de sniper");
            deAttacks.Add("Tueur de loup-garou");
            deAttacks.Add("Vitesse de la lumière noire");
            deAttacks.Add("Vol de vie");
            deAttacks.Add("Zone de mort psychique");
            // Gadgets
            deAttacks.Add("gadget de l'inspecteur");
            // Tokyo - enemys
            deAttacks.Add("Prototype de Rictus Orochi Mk 1.1");
            deAttacks.Add("Prototype de Rictus Orochi Mk 1.3");
            // Tokyo - enemy skills
            deAttacks.Add("Bain de sang");
            deAttacks.Add("Colonne de feu");
            deAttacks.Add("Coup de pied tombé");
            deAttacks.Add("Protocole de bouclier");
            deAttacks.Add("Système de réparation automatique");
            deAttacks.Add("Gaz de suppression");
            deAttacks.Add("Répulseur de combat");
            deAttacks.Add("Coup de griffes");
            deAttacks.Add("Consommation de ki");
            deAttacks.Add("Explosion de tentacules");
            deAttacks.Add("Gerbe de Souillure");
            deAttacks.Add("Pluie de sang");
            deAttacks.Add("Onde de souffle");
            deAttacks.Add("Frénésie de balles");
            deAttacks.Add("Dents de diamant");
            deAttacks.Add("Déluge de plomb");
            deAttacks.Add("Perturbation de la terre");
            deAttacks.Add("Brûlure de scorie");
            deAttacks.Add("Chant de l'archer");
            deAttacks.Add("Coulée de lave");
            deAttacks.Add("Coup de mandibules");
            deAttacks.Add("Lancer de grenade");
            deAttacks.Add("Pluie de putrescence");
            /*
            deAttacks.Add("");
            */

            // Tokyo - items and gadgets
            deAttacks.Add("Module de boucle de feedback améliorée");
            deAttacks.Add("Module de rajeunissement redirigé");
            deAttacks.Add("Module de boucle de feedback");
            deAttacks.Add("Module de dérivation améliorée");
            deAttacks.Add("Module de dérivation améliorée");
            deAttacks.Add("Module de dérivation pénétrante");
            deAttacks.Add("Module de dérivation directe");
            deAttacks.Add("Module de dérivation transposée");
            deAttacks.Add("Stimulant de vitalité");
		//SWL	
			//SWL - Attacks and Items
            deAttacks.Add("Âme de joueur");
			deAttacks.Add("Cacophonie de doppelgänger extradimensionnel");
			deAttacks.Add("Détonation de grenade");
			deAttacks.Add("Dissipation de doppelgänger extradimensionnel");
			deAttacks.Add("Enchaînement de coups");
			deAttacks.Add("Énigme de bien-être");
	        deAttacks.Add("Force de Gaia");		
			deAttacks.Add("Linceul de châtiment");
			deAttacks.Add("Potion de soins");
			deAttacks.Add("Regain de vigueur");
            deAttacks.Add("Soif de sang");
            deAttacks.Add("Surcroît de chi");
			deAttacks.Add("Suppression de la douleur");
            deAttacks.Add("Vague de chaleur");
			//SWL-Gadgets
            deAttacks.Add("Drone médical de modèle 3P-K");
			deAttacks.Add("Drone de purification de modèle UW-0");
			deAttacks.Add("Système d'inoculation de sédatif");

        // end of list
            StringBuilder deStringBuilder = new StringBuilder();
            foreach (var deAttack in deAttacks)
            {
                if (deStringBuilder.Length > 1)
                {
                    deStringBuilder.Append("|");
                }

                deStringBuilder.Append(deAttack);
            }

            string deString = deStringBuilder.ToString();

            damageLines = new List<Regex>();
            damageLines.Add(new Regex(@"^(?<actor>Votre)\s(?<attackName>.+)\sinflige\s(?<amount>[0-9]+)\spoints\sde\sdégâts\sà\sl\'(?<actee>.+)\.$", RegexOptions.Compiled));
            //damageLines.Add(new Regex(@"^(?<attackName>.+" + deString + @")\sde\s(?<actor>.+)\sinflige\s(?<amount>[0-9]+)\spoints\sde\sdégâts\sà\sl\'(?<actee>.+)\.$", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<attackName>.+" + deString + @")\sde\s(?<actor>.+)\sinflige\s(?<amount>[0-9]+)\spoints\sde\sdégâts\sà\sl\'(?<actee>.+)\.$"));
            damageLines.Add(new Regex(@"^(?<attackName>.+?)\sde\s(?<actor>.+)\sinflige\s(?<amount>[0-9]+)\spoints\sde\sdégâts\sà\sl\'(?<actee>.+)\.$", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?(?<actor>Vous)\stouchez\s(?<actee>.+)\savec\s(?<attackName>.+?)\s(?<damageType>\([^\)]+\)\s)(sur\sun\scritique\s)?et\slui\sinfligez\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled)); // SWL
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?(?<actor>.+)\stouche\s(?<actee>.+)\savec\s(?<attackName>.+?)\s(?<damageType>\([^\)]+\)\s)(sur\sun\scritique\s)?et\slui\sinflige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));  // SWL
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)(?<actor>Vous)\stouchez\s(?<damageType>\([^\)]+\)\s)(?<actee>.+)\savec\s(?<attackName>.+?)\s(sur\sun\scritique\s)?et\slui\sinfligez\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled)); // TSW
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)(?<actor>.+)\stouche\s(?<damageType>\([^\)]+\)\s)(?<actee>.+)\savec\s(?<attackName>.+?)\s(sur\sun\scritique\s)?et\slui\sinflige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));  // TSW
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?(?<actor>Votre)\s(?<attackName>.+)\stouche\s(?<damageType>\([^\)]+\)\s)(?<actee>.+)\set\slui\sinflige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?Le\spouvoir\s(?<attackName>" + deString + @")\sde\s(?<actor>.+)\stouche\s(?<actee>.+)\s(?<damageType>\([^\)]+\)\s)et\slui\sinflige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?"));
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?Le\spouvoir\s(?<attackName>" + deString + @")\sde\s(?<actor>.+)\s(?<actee>vous)\stouche\s(?<damageType>\([^\)]+\)\s)et\svous\sinflige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?"));
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?Le\spouvoir\s(?<attackName>.+?)\sde\s(?<actor>.+)\stouche\s(?<actee>.+)\s(?<damageType>\([^\)]+\)\s)et\slui\sinflige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?Le\spouvoir\s(?<attackName>.+?)\sde\s(?<actor>.+)\s(?<actee>vous)\stouche\s(?<damageType>\([^\)]+\)\s)et\svous\sinflige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\s(de\s)?(?<damageClass>.*?)\.(?<blockType>\s\([^\)]+\))?", RegexOptions.Compiled));
            damageLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?(?<actor>Votre)\spouvoir\s(?<attackName>.+?)\s(?<actee>vous)\stouche\set\svous\sinflige\s(?<amount>[0-9]+)\spoints\sde\sdégâts", RegexOptions.Compiled));

            //TSW Language Mixup Hack
            damageLines.Add(new Regex(@"^(?<actor>Votre|.*'s)\s(?<attackName>.+)\sdealt\s(?<amount>[0-9]+)\sdamage\sto\s(?<actee>.+)\.$", RegexOptions.Compiled));

            string hateSkills = "Provocation|Provocation de masse|Confusion|Confusion de masse|Fausse cible";
            damageLines.Add(new Regex(@"^(?<actor>.+)\s(avez|a)\sutilisé\sle\spouvoir\s(?<attackName>" + hateSkills + @")\savec\ssuccès\.", RegexOptions.Compiled));

            redirectLines = new List<Regex>();
            redirectLines.Add(new Regex(@"^(?<actor>Votre)\s(pouvoir\s)?(?<attackName>.+)\sredirige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\svers\s(?<actee>.+)\.$", RegexOptions.Compiled));
            redirectLines.Add(new Regex(@"^Le\spouvoir\s(?<attackName>" + deString + @")\sde\s(?<actor>.+)\sredirige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\svers\s(?<actee>.+)\.$"));
            redirectLines.Add(new Regex(@"^Le\spouvoir\s(?<attackName>.+?)\sde\s(?<actor>.+)\sredirige\s(?<amount>[0-9]+)\spoints?\sde\sdégâts\svers\s(?<actee>.+)\.$", RegexOptions.Compiled));

            healLines = new List<Regex>();
            healLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?(?<actor>Votre)\s(pouvoir\s)?(?<attackName>.+)\srend\s(?<amount>[0-9]+)\spoints?\sde\svie\sà\s(?<actee>.+)\.$", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?(?<actor>Votre)\s(pouvoir\s)?(?<attackName>.+)\s(?<actee>vous)\srend\s(?<amount>[0-9]+)\spoints?\sde\svie\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?(?<actor>.+)\srend\s(?<amount>[0-9]+)\spoints?\sde\svie\sà\s(?<actee>.+)\savec\s(?<attackName>.+)\.$", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?(?<actor>.+)\s(?<actee>vous)\srend\s(?<amount>[0-9]+)\spoints?\sde\svie\savec\s(?<attackName>.+)\.$", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?Le\spouvoir\s(?<attackName>" + deString + @")\sde\s(?<actor>.+)\srend\s(?<amount>[0-9]+)\spoints?\sde\svie\sà\s(?<actee>.+)\.$"));
            healLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?Le\spouvoir\s(?<attackName>.+?)\sde\s(?<actor>.+)\srend\s(?<amount>[0-9]+)\spoints?\sde\svie\sà\s(?<actee>.+)\.$", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?Le\spouvoir\s(?<attackName>" + deString + @")\sde\s(?<actor>.+)\s(?<actee>vous)\srend\s(?<amount>[0-9]+)\spoints?\sde\svie\."));
            healLines.Add(new Regex(@"^(?<crit>\(Critique\)\s)?Le\spouvoir\s(?<attackName>.+?)\sde\s(?<actor>.+)\s(?<actee>vous)\srend\s(?<amount>[0-9]+)\spoints?\sde\svie\.", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<attackName>" + deString + @")\sde\s(?<actor>.+)\srend\s(?<amount>[0-9]+)\spoints\sde\srésistance\sà\sl\'(?<actee>.+)\.$", RegexOptions.Compiled));
            healLines.Add(new Regex(@"^(?<attackName>.+?)\sde\s(?<actor>.+)\srend\s(?<amount>[0-9]+)\spoints\sde\srésistance\sà\sl\'(?<actee>.+)\.$", RegexOptions.Compiled));

            evadedLines = new List<Regex>();
            evadedLines.Add(new Regex(@"^(?<actee>.+)\sévite\sle\spouvoir\s(?<attackName>" + deString + @")\sde\s(?<actor>.+)\.$"));
            evadedLines.Add(new Regex(@"^(?<actee>.+)\sévite\sle\spouvoir\s(?<attackName>.+)\sde\s(?<actor>.+)\.$", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actee>Vous)\savez\sévité\s(?<attackName>" + deString + @"),\sde\s(?<actor>.+)\.$"));
            evadedLines.Add(new Regex(@"^(?<actee>Vous)\savez\sévité\s(?<attackName>.+),\sde\s(?<actor>.+)\.$", RegexOptions.Compiled));
            evadedLines.Add(new Regex(@"^(?<actee>.+)\sa\sévité\s(?<actor>votre)\s(?<attackName>.+)\.$", RegexOptions.Compiled));

            diedLines = new List<Regex>();
            diedLines.Add(new Regex(@"^(?<actor>.*)\sa\sperdu\sla\svie\.", RegexOptions.Compiled));

            buffLines = new List<Regex>();
            buffLines.Add(new Regex(@"^Vous\srecevez\sle\sbuff\s:\s(?<name>.*)$", RegexOptions.Compiled));
            buffLines.Add(new Regex(@"\svous\saffecte\savec\s(?<name>.*)$", RegexOptions.Compiled));

            buffStopLines = new List<Regex>();
            buffStopLines.Add(new Regex(@"^Buff\s(?<name>.*)\sterminé\.", RegexOptions.Compiled));

            aegisShieldLines = new List<Regex>();
            aegisShieldLines.Add(new Regex(@"AEGIS\s(?<aegis>\S+)\sde\s(?<actee>.+)", RegexOptions.Compiled));
            aegisShieldLines.Add(new Regex(@"^(?<actee>votre)\sAEGIS\s(?<aegis>.+)", RegexOptions.Compiled));

            //TSW Language Mixup Hack
            aegisShieldLines.Add(new Regex(@"^(?<actee>your)\s(?<aegis>.+)\sAEGIS", RegexOptions.Compiled));

            // Special damage types
            none = "none";
            Unknown = "inconnu";

            // Additional parsing strings
            Killing = "Mort";
            Miss = "Miss";
        }

        public static Boolean Init()
        {
            SetEnglish();
            return true;
        }
    }
}
