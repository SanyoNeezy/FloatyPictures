using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace FloatyButtons
{
    public static class DocumentManipulator
    {
        public static Word.Selection activeSelection;
        public static Word.Document activeDocument;

        /// <summary>
        /// Add a Table to your Document with the specified number of Rows and Columns
        /// </summary>
        /// <param name="numRows"></param>
        /// <param name="numColumns"></param>
        /// <returns></returns>
        public static Word.Table AddTable(int numRows, int numColumns)
        {
            activeDocument.Tables.Add(Range: activeSelection.Range, NumRows: numRows, NumColumns: numColumns,
            DefaultTableBehavior: Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior: Word.WdAutoFitBehavior.wdAutoFitFixed);

            Word.Table table = activeSelection.Tables[1];
            table.set_Style("Table Grid");
            table.ApplyStyleHeadingRows = true;
            table.ApplyStyleLastRow = false;
            table.ApplyStyleFirstColumn = true;
            table.ApplyStyleLastColumn = false;
            table.ApplyStyleRowBands = true;
            table.ApplyStyleColumnBands = false;

            return table;
        }

        /// <summary>
        /// gets and returns the selected image
        /// </summary>
        public static Word.InlineShape GetSelectedShape()
        {
            
            Word.InlineShape shape = null;
            try
            {
                if (!(activeSelection.Range.InlineShapes[1] is null))
                {
                    shape = activeSelection.Range.InlineShapes[1];
                    activeSelection.ChildShapeRange.WrapFormat.Type = Word.WdWrapType.wdWrapThrough;
                }
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                MessageBox.Show(e.Message);
            }

            return shape;
        }


        public static void FloatImages(bool onlySelected = false)
        {
            try
            {

                switch (activeSelection.Type)
                {
                    //Fall: Nur eine Inline Shape (also ein Bild) ausgewählt --> das ausgewählte Bild wird floaty
                    case Word.WdSelectionType.wdSelectionInlineShape:
                        activeSelection.ChildShapeRange.WrapFormat.Type = Word.WdWrapType.wdWrapThrough;
                        break;
                    //Fall: Auswahl mit keinen oder mehreren Bilder --> Alle ausgewählten Bilder werden floaty
                    case Word.WdSelectionType.wdSelectionNormal:
                        foreach (Word.InlineShape inlineShape in activeSelection.InlineShapes)
                        {
                            Word.Shape shape = inlineShape.ConvertToShape();
                            shape.WrapFormat.Type = Word.WdWrapType.wdWrapThrough;
                        }
                        break;
                    //Fall: Keine Auswahl  --> Alle Bilder im gesamten Dokument werden floaty
                    case Word.WdSelectionType.wdSelectionIP:
                        foreach (Word.InlineShape inlineShape in activeDocument.InlineShapes)
                        {
                            Word.Shape shape = inlineShape.ConvertToShape();
                            shape.WrapFormat.Type = Word.WdWrapType.wdWrapThrough;
                        }
                        break;
                }                 
                

            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                MessageBox.Show(e.Message);
            }


        }

        
    }



}
