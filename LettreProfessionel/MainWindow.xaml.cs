using LettreProfessionel.Models;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace LettreProfessionel
{
    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<LettreModel> lettres = [];

        public MainWindow()
        {
            InitializeComponent();
            ChargerLettres();
        }

        private void ChargerLettres()
        {
            string dossierLettres = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets/TemplateLettre");

            if (!Directory.Exists(dossierLettres))
            {
                MessageBox.Show($"Le dossier {dossierLettres} n'existe pas.");
                return;
            }

            string[] fichiersHtml = Directory.GetFiles(dossierLettres, "*.html");
            int compteur = 1;

            foreach (string fichier in fichiersHtml)
            {
                string titre = ExtraireTitreDepuisFichier(fichier);
                lettres.Add(new LettreModel
                {
                    Index = compteur++,
                    Titre = titre,
                    CheminFichier = fichier
                });
            }

            LettresList.ItemsSource = lettres;
        }

        private string ExtraireTitreDepuisFichier(string cheminFichier)
        {
            try
            {
                string contenu = File.ReadAllText(cheminFichier);
                var match = Regex.Match(contenu, @"<title>(.*?)</title>", RegexOptions.IgnoreCase);

                if (match.Success)
                    return match.Groups[1].Value.Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la lecture du fichier {cheminFichier} : {ex.Message}");
            }

            return Path.GetFileNameWithoutExtension(cheminFichier);
        }

        private async void LettresList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (LettresList.SelectedItem is not LettreModel lettre)
            {
                StatusText.Text = "Prêt";
                LetterInfoText.Text = "";
                return;
            }

            if (!File.Exists(lettre.CheminFichier))
            {
                StatusText.Text = "Fichier introuvable";
                LetterInfoText.Text = "";
                return;
            }

            // Mettre à jour le status bar
            StatusText.Text = lettre.Titre;          
            LetterInfoText.Text = lettre.CheminFichier; 

            // Afficher dans WebView2
            await HtmlViewer.EnsureCoreWebView2Async();
            string uri = "file:///" + lettre.CheminFichier.Replace("\\", "/");
            HtmlViewer.CoreWebView2.Navigate(uri);
        }

        //imprimer la lettre sélectionnée
        private async void Print_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 == null)
            {
                MessageBox.Show("Veuillez sélectionner une lettre à imprimer.");
                return;
            }

            await HtmlViewer.CoreWebView2.ExecuteScriptAsync("window.print()");
            StatusText.Text = "Impression lancée";
        }

        //Gestion des commandes d'édition de texte 
        private async void Bold_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('bold');");
        }

        private async void Italic_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('italic');");
        }

        private async void Underline_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('underline');");
        }

        private async void AlignLeft_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('justifyLeft');");
        }

        private async void AlignCenter_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('justifyCenter');");
        }

        private async void AlignRight_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('justifyRight');");
        }

        private async void AlignJustify_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('justifyFull');");
        }

        // Annuler
        private async void Undo_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('undo');");
        }

        // Refaire
        private async void Redo_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('redo');");
        }

        // Liste à puces
        private async void bulletList_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('insertUnorderedList');");
        }

        // Liste numérotée
        private async void numberedList_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
                await HtmlViewer.CoreWebView2.ExecuteScriptAsync("document.execCommand('insertOrderedList');");
        }

        // Lien (ouvre une boîte pour saisir l'URL)
        private async void lien_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlViewer.CoreWebView2 != null)
            {
                string url = Microsoft.VisualBasic.Interaction.InputBox("Entrez l'URL :", "Insérer un lien", "https://");
                if (!string.IsNullOrWhiteSpace(url))
                {
                    await HtmlViewer.CoreWebView2.ExecuteScriptAsync($"document.execCommand('createLink', false, '{url}');");
                }
            }
        }



    }
}
