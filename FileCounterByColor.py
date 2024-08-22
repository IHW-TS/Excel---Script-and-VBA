function main(workbook: ExcelScript.Workbook) {
    // Obtenir la feuille de calcul active.
    let selectedSheet = workbook.getActiveWorksheet();

    // Définir la plage à analyser (par exemple, A1:A350). 

    let range = selectedSheet.getRange("A1:A300");

    // Initialiser les compteurs pour chaque couleur.
    let redCount: number = 0;
    let yellowCount: number = 0;
    let blueCount: number = 0;
    let purpleCount: number = 0;
    let brownCount: number = 0;

    // Parcourir chaque cellule de la plage et compter les couleurs.
    let rowCount = range.getRowCount();
    let colCount = range.getColumnCount();

    for (let i = 0; i < rowCount; i++) {
        for (let j = 0; j < colCount; j++) {
            let cell = range.getCell(i, j);
            let color = cell.getFormat().getFill().getColor();

            if (color === "#FF0000") { // Rouge
                redCount++;
            } else if (color === "#FFFF00") { // Jaune
                yellowCount++;
            } else if (color === "#00B0F0") { // Bleu
                blueCount++;
            } else if (color === "#FF30F8") { // Violet
                purpleCount++;
            } else if (color === "#ED7D31") { // Marron
                brownCount++;
            }
        }
    }

    let totalOtherColors = yellowCount + blueCount + brownCount;

    // Afficher les résultats dans un tableau.
    //selectedSheet.getRange("A1").setValue("Type de Dossier");
    //selectedSheet.getRange("B1").setValue("Valeur");
    //selectedSheet.getRange("A2").setValue("MED envoyée");
    selectedSheet.getRange("B2").setValue(redCount);
    //selectedSheet.getRange("A3").setValue("MED à traité");
    selectedSheet.getRange("B3").setValue(yellowCount);
    //selectedSheet.getRange("A4").setValue("Manque un document");
    selectedSheet.getRange("B4").setValue(brownCount);
    //selectedSheet.getRange("A5").setValue("Siège fermé");
    selectedSheet.getRange("B5").setValue(blueCount);
    //selectedSheet.getRange("A6").setValue("MED à valider par SG ou PL");
    selectedSheet.getRange("B6").setValue(purpleCount);
    //selectedSheet.getRange("A7").setValue("Total autres couleurs");
    selectedSheet.getRange("B7").setValue(totalOtherColors);

}
