using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    class CreateExcell
    {
        public void OpenFiles( string fromNazwisko, string fromStanowisko, string toCreateFile )
        {
            try
            {
                LoadOptions loPassword = new LoadOptions( );
                loPassword.Password = passwordNazwisko;
                using ( Workbook wbNazwisko = new Workbook( fromNazwisko, loPassword ) )
                //using ( Workbook wbNazwisko = new Workbook( "C:\\Users\\agalanyuk\\Desktop\\Excell\\From_720.xls", loPassword ) )
                {
                    IsOpenNaz = true;
                    loPassword.Password = passwordStanowisko;
                    using ( Workbook wbStanowisco = new Workbook( fromStanowisko, loPassword ) )
                    //using ( Workbook wbStanowisco = new Workbook( "C:\\Users\\agalanyuk\\Desktop\\Excell\\From_720 Rozdzielnik wynagrodzeń OFP_2023_2024_CAŁOŚC_FK.xlsx", loPassword ) )
                    {
                        loPassword.Password = passwordCreateFile;
                        IsOpenStan = true;

                        using ( Workbook wbCreateFile = new Workbook( toCreateFile, loPassword ) )
                        //using ( Workbook wbCreateFile = new Workbook( "C:\\Users\\agalanyuk\\Desktop\\Excell\\To_MPK 720 OFP PŁACE TYCHY 2023.XLS", loPassword ) )
                        {
                            IsOpenCreateFile = true;
                            path = toCreateFile;
                            Console.WriteLine( "All file are sucsses opened " );

                            if ( !CreateNewSheet( wbNazwisko, wbStanowisco, wbCreateFile ) )
                            {
                                throw new Exception( "Something wrong" );
                            }
                        }
                    }

                }
            }
            catch ( Exception ex )
            {

                MessageBox.Show( ex.Message );
            }
        }

        private bool CreateNewSheet( in Workbook wbNazwisko, in Workbook wbStanowisco, Workbook wbCreateFile )
        {
            bool result = true;
            Int32 index = wbCreateFile.Worksheets.Add( );
            Worksheet worksheet = wbCreateFile.Worksheets[index];

            result = CheckNazwisko( wbNazwisko.Worksheets[0], wbStanowisco.Worksheets[wbStanowisco.Worksheets.Count - 2] );
            if ( result )
            {
                if ( CreateTemplateOrder( worksheet, wbCreateFile ) < 0 )
                {
                    Console.WriteLine( "New report template is incorrect !!!" );
                    result = false;
                }

                if ( result )
                {
                    result = CopyNazwiskoData( wbNazwisko, worksheet );
                    if ( result )
                    {
                        result = CopyStanowiskoData( wbStanowisco, worksheet );
                        if ( result )
                        {
                            Workbook workbookNew = new Workbook( );
                            workbookNew.Worksheets[0].Copy( worksheet );
                            workbookNew.Save( "C:\\Users\\agalanyuk\\Desktop\\Excell\\new_report\\new_report.xls" );
                            MessageBox.Show( "New worksheet save is correct" );
                        }
                        else
                        {
                            MessageBox.Show( "Error CopyStanowiskoData" );
                        }
                    }
                    else
                    {
                        MessageBox.Show( "Error CopyNazwiskoData" );
                    }
                }
                else
                {
                    MessageBox.Show( "Error CreateTemplateOrder" );
                }
            }
            else
            {
                MessageBox.Show( "Nazwisko list are different" );
            }

            return result;
        }

        private bool CheckNazwisko( Worksheet wbNazwisko, Worksheet wbStanowisco )
        {
            bool result = true;
            String nazwiskoTitle = "nazwisko, imię";
            String stanowiskoTitle = "Nazwisko i imię";
            List<Index> inNazwisko = new List<Index>( );
            List<Index> inStanowisko = new List<Index>( );
            Int32 shift = 0;

            copyRowsCount = CountCopyRows( wbNazwisko );

            if ( !FindIndex( inNazwisko, nazwiskoTitle, wbNazwisko ) )
            {
                MessageBox.Show( "Don't find Nazwisko" );
            }

            if ( !FindIndex( inStanowisko, stanowiskoTitle, wbStanowisco ) )
            {
                MessageBox.Show( "Don't find Stanowisko" );
            }

            shift = OffsetSearh( wbStanowisco, copyRowsCount, inStanowisko[0] );

            //Console.WriteLine( $"{wbNazwisko.Cells[inNazwisko[0].Row + 1, inNazwisko[0].Column].Value} / " +
            //    $"{wbStanowisco.Cells[inStanowisko[0].Row + shift, inStanowisko[0].Column].Value}" );

            for ( int i = 1; i < copyRowsCount; i++ )
            {
                if ( !wbNazwisko.Cells[inNazwisko[0].Row + i, inNazwisko[0].Column].Value.ToString( ).
                    Equals( wbStanowisco.Cells[inStanowisko[0].Row + shift + i, inStanowisko[0].Column].Value.ToString( ) ) )
                {
                    MessageBox.Show( $"{wbNazwisko.Cells[inNazwisko[0].Row + i, inNazwisko[0].Column].Value.ToString( )} " +
                        $"/ {wbStanowisco.Cells[inStanowisko[0].Row + shift + i, inStanowisko[0].Column].Value.ToString( )}" );
                    result = false;
                    break;
                }
            }

            return result;
        }

        private Int32 OffsetSearh( in Worksheet wbStanowisco, in int copyRowsCount, in Index index )
        {
            Int32 shift = 0;

            Console.WriteLine( $"title: {wbStanowisco.Cells[index.Row, index.Column].Value}" );


            for ( int i = 1; i < copyRowsCount; ++i )
            {
                if ( wbStanowisco.Cells[index.Row + i, index.Column].Value == null )
                {
                    shift++;
                }
                else
                {
                    Console.WriteLine( $"off set: {wbStanowisco.Cells[index.Row + i, index.Column].Value}" );
                    break;
                }
            }

            return shift;
        }

        private Int32 CreateTemplateOrder( in Worksheet worksheet, Workbook wbCreateFile )
        {
            Int32 index = -1;
            string subStr = "TYCHY";

            for ( int i = wbCreateFile.Worksheets.Count - 1; i > 1; i-- )
            {
                if ( wbCreateFile.Worksheets[i].Name.Contains( subStr ) )
                {
                    MessageBox.Show( wbCreateFile.Worksheets[i].Name );
                    index = i;
                    break;
                }
            }

            worksheet.Copy( wbCreateFile.Worksheets[index] );
            worksheet.Name = "new report 1";
            return index;
        }

        private bool CopyNazwiskoData( in Workbook workbook, Worksheet worksheet )
        {
            bool result = true;
            List<Index> fromDataIndexes = new List<Index>( );
            List<string> fromData = new List<string>
            {
                "nazwisko, imię",
                "Wynagrodzenie Brutto",
            };

            foreach ( var str in fromData )
            {
                if ( !FindIndex( fromDataIndexes, str, workbook.Worksheets[0] ) )
                {
                    MessageBox.Show( $"Don't find title: {str}" );
                    return false;
                }
            }

            List<Index> toDataIndexes = new List<Index>( );
            List<string> toData = new List<string>

            {
                "nazwisko, imię",
                "Wynagrodz. Brutto",
                " ZUS + FP"
            };

            foreach ( var str in toData )
            {
                if ( !FindIndex( toDataIndexes, str, worksheet ) )
                {
                    MessageBox.Show( $"Don't find title: {str}" );
                    return false;
                }
            }

            //copyRowsCount = CountCopyRows( workbook );

            Console.WriteLine( $"count copy rows: {copyRowsCount}" );

            result = copyFileNazwisko( workbook.Worksheets[0], worksheet, copyRowsCount, fromDataIndexes, toDataIndexes );

            return result;
        }

        private bool CopyStanowiskoData( in Workbook wbStanowisco, Worksheet worksheet )
        {
            bool result = true;
            List<Index> fromDataIndex = new List<Index>( );
            List<string> titleData = new List<string>
            {
                "Administracja",
                "BSI",
                "BSI IPET",
                "T",
                "T IPET",
                "LOM",
                "LOM IPET",
                "LP"
            };
            Int32 countData = 8;

            foreach ( var data in titleData )
            {
                if ( !FindIndex( fromDataIndex, data, wbStanowisco.Worksheets[1] ) )
                {
                    MessageBox.Show( $"CopyStanowiskoData:: don't fined data title: {data}" );
                    result = false;
                }
            }

            if ( fromDataIndex.Count == countData && result )
            {
                Console.WriteLine( "All to data were found" );
            }

            if ( result )
            {
                List<Index> toDataIndex = new List<Index>( );
                List<string> toTitleData = new List<string>
                {
                    "Stanowisko",
                    "BSI",
                    "BSI IPET",
                    "T",
                    "T   IPET",
                    "LOM",
                    "LOM IPET",
                    "LP"
                };

                foreach ( var data in toTitleData )
                {
                    if ( !FindIndex( toDataIndex, data, worksheet ) )
                    {
                        MessageBox.Show( $"CopyStanowiskoData:: don't fined data title: {data}" );
                        result = false;
                    }
                }

                if ( toDataIndex.Count == countData && result )
                {
                    Console.WriteLine( "All to data were found" );
                }

                if ( result )
                {
                    result = copyStanowiskoFile( wbStanowisco.Worksheets[1], worksheet, fromDataIndex, toDataIndex );
                }
            }

            return result;
        }

        private Int32 CountCopyRows( in Worksheet worksheet )
        {
            Int32 countRows = -1;
            for ( int i = 0; i < worksheet.Cells.Rows.Count - 1; i++ )
            {
                if ( null == worksheet.Cells[i + 2, 0].Value )
                {
                    countRows = i;
                    break;
                }
            }
            return countRows;
        }

        private bool FindIndex( List<Index> dataIndexes, string findObgect, Worksheet worksheet )
        {
            bool result = false;

            Int32 rowCount = worksheet.Cells.Rows.Count;
            Int32 columnCount = worksheet.Cells.Columns.Count;

            for ( int column = 0; column <= columnCount; column++ )
            {
                for ( int row = 0; row <= rowCount; row++ )
                {
                    //if ( column == 6 )
                    //{
                    //    Console.WriteLine( );
                    //}


                    if ( worksheet.Cells[row, column].Value != null )
                    {
                        if ( worksheet.Cells[row, column].Value.ToString( ).Equals( findObgect ) )
                        {
                            dataIndexes.Add( new Index( row, column ) );
                            result = true;
                            column = worksheet.Cells.Columns.Count;
                            break;
                        }
                    }
                }
            }
            return result;
        }

        private bool copyFileNazwisko( in Worksheet fromWorksheet, Worksheet toWorksheet,
                                                in Int32 rowCount, in List<Index> fromData,
                                               in List<Index> toData )
        {
            bool result = true;
            bool isInsert = false;
            Int32 columnNazwisko = 1;
            Int32 nazwiskoValue = fromData[0].Column;
            Int32 wynagValue = fromData[1].Column;
            Int32 deleteRows = 0;

            for ( int i = 1; i <= rowCount; i++ )
            {
                if ( toWorksheet.Cells[i + 1, 1].Value == null && toWorksheet.Cells[i + 2, 1].Value == null )
                {
                    isInsert = true;
                    toWorksheet.Cells.InsertRows( i + 1, columnNazwisko );
                    toWorksheet.Cells.CopyRow( toWorksheet.Cells, i, i + 1 );
                    Console.WriteLine( $" copyFileNazwisko:: insert row: {fromWorksheet.Cells[i + 1, nazwiskoValue].Value.ToString( )}" );
                    toWorksheet.Cells[i + 1, 0].PutValue( fromWorksheet.Cells[i, nazwiskoValue].Value );
                }
                else
                {
                    for ( int index = 0; index < toData.Count; ++index )
                    {
                        if ( toWorksheet.Cells[0, toData[index].Column].Value.ToString( ).Equals( " ZUS + FP" ) )
                        {
                            toWorksheet.Cells[i, toData[index].Column].PutValue(
                                SummData( ( double ) fromWorksheet.Cells[fromData[index - 1].Row + i, wynagValue + 1].Value,
                        ( double ) fromWorksheet.Cells[fromData[index - 1].Row + i, wynagValue + 2].Value ) );
                        }
                        else
                        {
                            toWorksheet.Cells[toData[index].Row + i, toData[index].Column].PutValue(
                                fromWorksheet.Cells[fromData[index].Row + i, fromData[index].Column].Value );
                        }
                    }
                }
            }

            if ( toWorksheet.Cells[rowCount + 1, 1].Value != null && !isInsert )
            {
                for ( int i = rowCount + 1; toWorksheet.Cells[i, 1].Value != null; ++i )
                {
                    Console.WriteLine( $" copyFileNazwisko:: Delete row: {toWorksheet.Cells[i, 1].Value.ToString( )}" );
                    deleteRows += 1;
                }
            }

            if ( deleteRows > 0 )
            {
                Console.WriteLine( $" copyFileNazwisko:: Delete rows: {deleteRows}" );
                toWorksheet.Cells.DeleteRows( rowCount + 1, deleteRows );
            }

            result = CheckCopyNazwisko( fromWorksheet, toWorksheet, fromData, toData );

            return result;
        }

        private bool CheckCopyNazwisko( in Worksheet originalWorksheet, in Worksheet copyWorksheet, in List<Index> originalData,
                                               in List<Index> copyData )
        {
            bool result = true;

            for ( int i = 1; i <= copyRowsCount; i++ )
            {
                for ( int j = 0; j < originalData.Count; ++j )
                {
                    //Console.WriteLine( $"{originalWorksheet.Cells[originalData[j].Row + i, originalData[j].Column].Value.ToString( )} / " +
                    //    $"{copyWorksheet.Cells[copyData[j].Row + i, copyData[j].Column].Value.ToString( )}");
                    if ( !originalWorksheet.Cells[originalData[j].Row + i, originalData[j].Column].Value.ToString( ).
                        Equals( copyWorksheet.Cells[copyData[j].Row + i, copyData[j].Column].Value.ToString( ) ) )
                    {
                        MessageBox.Show( $"{originalWorksheet.Cells[originalData[j].Row + i, originalData[j].Column].Value.ToString( )} \n  " +
                            $"{copyWorksheet.Cells[copyData[j].Row + i, copyData[j].Column].Value.ToString( )}" );
                        result = false;
                        break;
                    }
                }
            }

            return result;
        }

        private bool copyStanowiskoFile( in Worksheet fromWorksheet, Worksheet toWorksheet,
                                                 in List<Index> fromData, in List<Index> toData )
        {
            bool result = true;
            Int32 shift = 0;

            shift = OffsetSearh( fromWorksheet, copyRowsCount, fromData[1] );

            Console.WriteLine( $"copyStanowiskoFile:: shift: {shift}" );

            for ( int i = 1; i <= copyRowsCount; i++ )
            {
                for ( int index = 0; index < toData.Count; ++index )
                {
                    if ( index == 0 )
                    {
                        toWorksheet.Cells[toData[index].Row + i, toData[index].Column].
                            PutValue( fromWorksheet.Cells[fromData[index].Row + i, fromData[index].Column].Value );
                    }
                    else
                    {
                        //Console.WriteLine( $"name: {toWorksheet.Cells[toData[index].Row + i, 1].Value}" +
                        //    $" to : {toWorksheet.Cells[toData[index].Row + i, toData[index].Column - 1].Value} " +
                        //    $"from {fromWorksheet.Cells[fromData[index].Row + shift + i, fromData[index].Column].Value}" );

                        toWorksheet.Cells[toData[index].Row + i, toData[index].Column - 1].
                            PutValue( fromWorksheet.Cells[fromData[index].Row + shift + i, fromData[index].Column].Value );
                    }
                }
            }

            result = CheckCopyStanowisko( fromWorksheet, toWorksheet, fromData, toData, shift );

            return result;
        }

        private bool CheckCopyStanowisko( Worksheet originalWorksheet, Worksheet copyWorksheet, List<Index> originalData, List<Index> copyData, int shift )
        {
            bool result = true;

            for ( int i = 1; i <= copyRowsCount; i++ )
            {
                for ( int index = 0; index < originalData.Count; ++index )
                {
                    if ( index == 0 )
                    {
                        //Console.WriteLine( $"{originalWorksheet.Cells[originalData[index].Row + i, originalData[index].Column].Value.ToString( )} == " +
                        //    $"{copyWorksheet.Cells[copyData[index].Row + i, copyData[index].Column].Value.ToString( )}" );

                        if ( originalWorksheet.Cells[originalData[index].Row + i, originalData[index].Column].Value != null &&
                        !originalWorksheet.Cells[originalData[index].Row + i, originalData[index].Column].Value.ToString( ).
                            Equals( copyWorksheet.Cells[copyData[index].Row + i, copyData[index].Column].Value.ToString( ) ) ) 
                        {
                            result = false;
                            MessageBox.Show( $" or: {originalWorksheet.Cells[originalData[index].Row + i, originalData[index].Column].Value.ToString( )} /n " +
                                $"{copyWorksheet.Cells[copyData[index].Row + i, copyData[index].Column].Value.ToString( )}" );
                            break;
                        }
                    }
                    else
                    {
                        if ( !originalWorksheet.Cells[originalData[index].Row + shift + i, originalData[index].Column].Value.ToString( ).
                            Equals( copyWorksheet.Cells[copyData[index].Row + i, copyData[index].Column - 1].Value.ToString( ) ) )
                        {
                            result = false;
                            MessageBox.Show( $" or: {originalWorksheet.Cells[originalData[index].Row + shift + i, originalData[index].Column].Value.ToString( )} /n " +
                                $"{copyWorksheet.Cells[copyData[index].Row + i, copyData[index].Column].Value.ToString( )}" );
                            break;
                        }
                    }
                }
            }

            return result;
        }

        private double SummData( in double lValue, in double rValue )
        {
            return lValue + rValue;
        }

        public bool IsOpenNaz { get; set; }
        public bool IsOpenStan { get; set; }
        public bool IsOpenCreateFile { get; set; }

        public bool isPasswordNazwisko { get; set; }
        public bool isPasswordStanowisko { get; set; }
        public bool isPasswordCreateFile { get; set; }

        public string passwordNazwisko { get; set; }
        public string passwordStanowisko { get; set; }
        public string passwordCreateFile { get; set; }

        public string path { get; set; }
        public Int32 copyRowsCount { get; set; }
    }
    struct Index
    {
        public Index( Int32 row, Int32 column )
        {
            Row = row;
            Column = column;
        }
        public Int32 Row { get; set; }
        public Int32 Column { get; set; }
    }
}
