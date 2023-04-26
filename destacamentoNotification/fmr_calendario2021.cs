using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;
using criapLibrary.types;
using criapLibrary;
using DevExpress.Utils.DragDrop;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Drawing;

namespace CRIAPContratos
{
    public partial class Fmr_calendario2021 : Form
    {
        public string actualCode = null, actualRef = null;
        public DateTime especial_1 = new DateTime(), especial_2 = new DateTime(), extraordinaria = new DateTime(), extraordinariaINS = new DateTime();
        public List<Avaliacao> avaliacoes = new List<Avaliacao>(), avaliacoesativ = new List<Avaliacao>();
        public List<AvaliacaoCalculo> avaliacaocalculo = new List<AvaliacaoCalculo>();
        public List<Metolologias> metolologias = new List<Metolologias>();
        public List<ModulosSelect> modulosselect = new List<ModulosSelect>();
        public List<FeriadosHT> feriadosHT = new List<FeriadosHT>();
        Form1 mainForm = (Form1)Application.OpenForms["Form1"];
        public List<objCalendario.modulo> modulosHT = new List<objCalendario.modulo>(), modulosHTAll = new List<objCalendario.modulo>(), modulosHTSessoes = new List<objCalendario.modulo>();
        public List<library.types.Sessao> sessoesHT = new List<library.types.Sessao>();
        public DataTable tableCA = new DataTable();
        List<Rectangle> lines = new List<Rectangle>();

        public class Avaliacao
        {
            public int ID { get; set; }
            public string Codigo_Curso { get; set; }
            public string RefAcao { get; set; }
            public string MomentoAv { get; set; }
            public string Modulo { get; set; }
            public int ModuloID { get; set; }
            public int Ordem { get; set; }
            public float PesoCF { get; set; }
            public string Metodologia { get; set; }
            public string PesoAv { get; set; }
            public DateTime? DataI_EN { get; set; }
            public DateTime? DataF_EN { get; set; }
            public DateTime? DataI_ER { get; set; }
            public DateTime? DataF_ER { get; set; }
            public DateTime? DataI_EE { get; set; }
            public DateTime? DataF_EE { get; set; }
            public DateTime? DataI_EXT { get; set; }
            public DateTime? DataF_EXT { get; set; }
            public DateTime? Data_ER { get; set; }
            public DateTime? Data_EE { get; set; }
            public DateTime? Data_EXT { get; set; }
            public int AvaliacaoN { get; set; }
            public int Atividade { get; set; }
            public bool select { get; set; }
        }
        public class AvaliacaoCalculo
        {
            public int AvaliacaoN { get; set; }
            public DateTime DataI_ER { get; set; }
            public DateTime DataF_ER { get; set; }
        }
        public class Metolologias
        {
            public int Metolologia_ID { get; set; }
            public string Descricao { get; set; }
        }
        public class ModulosSelect
        {
            public int Modulo_ID { get; set; }
            public string Descricao { get; set; }
        }
        public class FeriadosHT
        {
            public DateTime Data { get; set; }
        }
        public Fmr_calendario2021()
        {
            InitializeComponent();
            GridView view = gridControl2.MainView as GridView;
            view.OptionsBehavior.Editable = false;
            gridControl2.Paint += GridControl2_Paint;
            HandleBehaviorDragDropEvents();
            actualCode = DB.selectCurso.Codigo_Curso;
            actualRef = DB.selectCurso.Ref_Accao;
        }
        public void HandleBehaviorDragDropEvents()
        {
            DragDropBehavior gridControlBehavior = behaviorManager1.GetBehavior<DragDropBehavior>(gridView2);
            gridControlBehavior.DragDrop += Behavior_DragDrop;
            gridControlBehavior.DragOver += Behavior_DragOver;
        }
        private void Behavior_DragOver(object sender, DragOverEventArgs e)
        {
            DragOverGridEventArgs args = DragOverGridEventArgs.GetDragOverGridEventArgs(e);
            e.InsertType = args.InsertType;
            e.InsertIndicatorLocation = args.InsertIndicatorLocation;
            e.Action = args.Action;
            Cursor.Current = args.Cursor;
            args.Handled = true;
        }
        private void Behavior_DragDrop(object sender, DragDropEventArgs e)
        {
            GridView targetGrid = e.Target as GridView;
            GridView sourceGrid = e.Source as GridView;
            if (e.Action == DragDropActions.None || targetGrid != sourceGrid)
                return;
            DataTable sourceTable = sourceGrid.GridControl.DataSource as DataTable;
            Point hitPoint = targetGrid.GridControl.PointToClient(Cursor.Position);
            GridHitInfo hitInfo = targetGrid.CalcHitInfo(hitPoint);

            int[] sourceHandles = e.GetData<int[]>();
            int targetRowHandle = hitInfo.RowHandle;
            int targetRowIndex = targetGrid.GetDataSourceRowIndex(targetRowHandle);
            List<DataRow> draggedRows = new List<DataRow>();

            foreach (int sourceHandle in sourceHandles)
            {
                int oldRowIndex = sourceGrid.GetDataSourceRowIndex(sourceHandle);
                DataRow oldRow = sourceTable.Rows[oldRowIndex];
                draggedRows.Add(oldRow);
            }

            int newRowIndex;

            switch (e.InsertType)
            {
                case InsertType.Before:
                    newRowIndex = targetRowIndex > sourceHandles[sourceHandles.Length - 1] ? targetRowIndex - 1 : targetRowIndex;
                    for (int i = draggedRows.Count - 1; i >= 0; i--)
                    {
                        DataRow oldRow = draggedRows[i];
                        DataRow newRow = sourceTable.NewRow();
                        newRow.ItemArray = oldRow.ItemArray;
                        sourceTable.Rows.Remove(oldRow);
                        sourceTable.Rows.InsertAt(newRow, newRowIndex);
                    }
                    break;
                case InsertType.After:
                    newRowIndex = targetRowIndex < sourceHandles[0] ? targetRowIndex + 1 : targetRowIndex;
                    for (int i = 0; i < draggedRows.Count; i++)
                    {
                        DataRow oldRow = draggedRows[i];
                        DataRow newRow = sourceTable.NewRow();
                        newRow.ItemArray = oldRow.ItemArray;
                        sourceTable.Rows.Remove(oldRow);
                        sourceTable.Rows.InsertAt(newRow, newRowIndex);
                    }
                    break;
                default:
                    newRowIndex = -1;
                    break;
            }
            int insertedIndex = targetGrid.GetRowHandle(newRowIndex);
            targetGrid.FocusedRowHandle = insertedIndex;
            targetGrid.SelectRow(targetGrid.FocusedRowHandle);
        }
        private void GridControl2_Paint(object sender, PaintEventArgs e)
        {
            GridControl grid = (GridControl)sender;
            GridView view = (GridView)grid.MainView;
            UpdateLines(view);
            foreach (Rectangle item in lines)
            {
                e.Graphics.FillRectangle(Brushes.Black, item);
            }
        }
        private void UpdateLines(GridView view)
        {
            lines.Clear();
            GridViewInfo viewInfo = view.GetViewInfo() as GridViewInfo;
            List<GridDataRowInfo> visibleRows = viewInfo.RowsInfo.OfType<GridDataRowInfo>().ToList();
            int prevDate = 1, rowAv;
            GridDataRowInfo currentRow;

            for (int i = 0; i < visibleRows.Count; i++)
            {
                currentRow = visibleRows[i];
                rowAv = Convert.ToInt32(view.GetRowCellValue(currentRow.RowHandle, "AvaliacaoN"));
                if (i > 0 && prevDate.ToString() != rowAv.ToString())
                {
                    lines.Add(new Rectangle(currentRow.Bounds.Left, currentRow.Bounds.Top - 1, currentRow.Bounds.Width, 2));
                }
                prevDate = rowAv;
            }
        }
        private void SimpleButton1_Click(object sender, EventArgs e)
        {
            mainForm.barButtonItem10.Enabled = true;
            mainForm.barButtonItem11.Enabled = false;
            mainForm.barButtonItem12.Enabled = false;
            mainForm.barButtonItem13.Enabled = false;
            Close();
        }
        private void Fmr_calendario2021_Load(object sender, EventArgs e)
        {
            GetFeriadosHT();
            GetModulosHT(actualCode, actualRef);
            GetSessoes(actualRef);
            modulosHT.Reverse();
            int nivel = 1;

            foreach (var item in modulosHT)
            {
                item.Ordem = nivel++;
            }

            modulosHTAll.Clear();
            List<DataRow> dataList = sv.getModulosHT(actualCode, actualRef, Connects.HTlocalConnect.Conn);
            if (dataList.Count != 0)
            {
                for (int i = 0; i < dataList.Count; i++)
                {
                    var obj = new objCalendario.modulo()
                    {
                        Nivel = dataList[i][1].ToString(),
                        Descricao = dataList[i][2].ToString(),
                        N_Horas = float.Parse(dataList[i][3].ToString()),
                        Peso_Aval = float.Parse(dataList[i][4].ToString()),
                        Data_Inicio = (DateTime)(dataList[i][6]),
                        Data_Fim = (DateTime)(dataList[i][7]),
                        Modulo_ID = (int)dataList[i][0],
                        Ordem = -1,
                        Teste_ID = -1,
                        check = false
                    };
                    modulosHTAll.Add(obj);
                }
            }

            //Apaga sessoes no mesmo dia em horas diferentes
            modulosHTSessoes.Clear();
            modulosHTAll.Reverse();
            foreach (var item2 in modulosHTAll)
            {
                int existe = modulosHTSessoes.Where(x => x.Nivel == item2.Nivel && x.Data_Inicio.Date == item2.Data_Inicio.Date).Count();
                if (existe == 0) modulosHTSessoes.Add(item2);
            }

            foreach (var item in modulosHT)
            {
                foreach (var item2 in modulosHTSessoes.Where(x => x.Nivel == item.Nivel))
                {
                    item2.Ordem = item.Ordem;
                }
            }

            int linha = 1;
            foreach (var item in modulosHTSessoes)
            {
                item.Teste_ID = linha++; //utilizado para identificar a posição do item
            }

            Carrega_calendario();
            Le_Modalidades();
            modulos.Items.Clear();
            for (int i = 0; i < modulosHT.Count; i++)
                modulos.Items.Add(modulosHT[i].Descricao);
        }
        public void CalendarioDTP()
        {
            GetFeriadosHT();
            GetModulosHT(actualCode, actualRef);
            GetSessoes(actualRef);
            modulosHT.Reverse();
        }
        public void Modulosclearandfill()
        {
            modulos.Items.Clear();
            for (int i = 0; i < modulosHT.Count; i++)
            {
                modulos.Items.Add(modulosHT[i].Descricao);
            }
        }
        public void Le_Modalidades()
        {
            metolologias.Clear();
            metodologia.Items.Clear();
            Connects.localConnect.ConnInit();
            string subQuery = "SELECT Modalidade_ID, Descricao FROM Modalidades where Modalidade_ID <> 5";
            SqlDataAdapter adapter = new SqlDataAdapter(subQuery, Connects.localConnect.Conn);
            DataTable subData = new DataTable();
            adapter.Fill(subData);
            List<DataRow> dataList = subData.AsEnumerable().ToList();

            if (dataList.Count != 0)
            {
                for (int i = 0; i < dataList.Count; i++)
                {
                    var obj = new Metolologias()
                    {
                        Metolologia_ID = (int)dataList[i][0],
                        Descricao = dataList[i][1].ToString()
                    };
                    metolologias.Add(obj);
                }
            }
            for (int i = 0; i < metolologias.Count; i++)
                metodologia.Items.Add(metolologias[i].Descricao);
            Connects.localConnect.ConnEnd();
        }
        public void Carrega_calendario()
        {
            avaliacoes.Clear();
            ID.Clear();
            //ModuloID.Clear();
            //pesocf.Clear();
            gridControl2.DataSource = null;
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Codigo_Curso");
            table.Columns.Add("RefAcao");
            table.Columns.Add("MomentoAv");
            table.Columns.Add("Modulo");
            table.Columns.Add("ModuloID");
            table.Columns.Add("Ordem");
            table.Columns.Add("PesoCF");
            table.Columns.Add("Metodologia");
            table.Columns.Add("PesoAv");
            table.Columns.Add("DataI_EN");
            table.Columns.Add("DataF_EN");
            table.Columns.Add("DataI_ER");
            table.Columns.Add("DataF_ER");
            table.Columns.Add("DataI_EE");
            table.Columns.Add("DataF_EE");
            table.Columns.Add("DataI_EXT");
            table.Columns.Add("DataF_EXT");
            table.Columns.Add("Data_ER");
            table.Columns.Add("Data_EE");
            table.Columns.Add("Data_EXT");
            table.Columns.Add("AvaliacaoN");

            string subQuery = "SELECT * FROM CalendarioAV WHERE Codigo_Curso = '" + actualCode + "' AND REFACAO = '" + actualRef + "' ORDER BY ORDEM";
            SqlDataAdapter adapter = new SqlDataAdapter(subQuery, Connects.localConnect.Conn);
            DataTable subData = new DataTable(); adapter.Fill(subData);
            List<DataRow> avaliacoesguardadas = subData.AsEnumerable().ToList();
            Connects.HTlocalConnect.ConnEnd();

            if (avaliacoesguardadas.Count > 0)
            {
                for (int j = 0; j < avaliacoesguardadas.Count; j++)
                {
                    table.Rows.Add(new object[]
                    {
                        int.Parse(avaliacoesguardadas[j][0].ToString()),
                        actualCode,
                        actualRef,
                        avaliacoesguardadas[j][2].ToString(),
                        int.Parse(avaliacoesguardadas[j][3].ToString()) == 0 ? "Época Extraordinária" : modulosHT.Where(x => x.Modulo_ID == int.Parse(avaliacoesguardadas[j][3].ToString())).First().Descricao,
                        int.Parse(avaliacoesguardadas[j][3].ToString()),
                        int.Parse(avaliacoesguardadas[j][4].ToString()),
                        float.Parse(avaliacoesguardadas[j][5].ToString()),
                        avaliacoesguardadas[j][6].ToString(),
                        avaliacoesguardadas[j][7].ToString(),
                        avaliacoesguardadas[j][8].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][8].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][9].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][9].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][10].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][10].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][11].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][11].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][12].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][12].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][13].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][13].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][14].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][14].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][15].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][15].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][18].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][18].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][19].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][19].ToString()).ToShortDateString(),
                        avaliacoesguardadas[j][20].ToString() == "" ? "" : DateTime.Parse(avaliacoesguardadas[j][20].ToString()).ToShortDateString(),
                        int.Parse(avaliacoesguardadas[j][16].ToString()),
                    });
                }

                GridView view = gridControl2.MainView as GridView;
                gridControl2.DataSource = table;
                tableCA = table;
                view.OptionsBehavior.Editable = false;
                label1.Text = DB.selectCurso.Descricao + "_" + actualRef + " // CALENDÁRIO DE AVALIAÇÃO EDITADO";
            }
            else
            {
                Lista();
            }
        }
        public void GetModulosHT(string codCurso, string refCurso)
        {
            modulosHT.Clear();
            List<DataRow> dataList = sv.getModulosHT(codCurso, refCurso, Connects.HTlocalConnect.Conn);

            if (dataList.Count != 0)
            {
                for (int i = 0; i < dataList.Count; i++)
                {
                    var obj = new objCalendario.modulo()
                    {
                        Nivel = dataList[i][1].ToString(),
                        Descricao = dataList[i][2].ToString(),
                        N_Horas = float.Parse(dataList[i][3].ToString()),
                        Peso_Aval = float.Parse(dataList[i][4].ToString()),
                        Data_Inicio = (DateTime)(dataList[i][6]),
                        Data_Fim = (DateTime)(dataList[i][7]),
                        Modulo_ID = (int)dataList[i][0],
                        Ordem = -1,
                        Teste_ID = -1,
                        check = false
                    };
                    modulosHT.Add(obj);
                }

                int j = 0;
                while (j < modulosHT.Count - 1)
                {
                    for (int k = j; k < modulosHT.Count - 1; k++)
                    {
                        if (modulosHT[j].Nivel == modulosHT[k + 1].Nivel)
                        {
                            modulosHT[j].Data_Inicio = modulosHT[k + 1].Data_Inicio;
                            modulosHT.RemoveAt(k + 1);
                            k--;
                        }
                    }
                    j++;
                }
            }
        }
        private void GetSessoes(string refCurso)
        {
            sessoesHT.Clear();
            List<DataRow> dataList = library.Rh.Get_Sessoes(Connects.HTlocalConnect.Conn, refCurso);

            if (dataList.Count != 0)
            {
                for (int i = 0; i < dataList.Count; i++)
                {
                    var obj = new library.types.Sessao()
                    {
                        rowid_Sessao = dataList[i][0].ToString(),
                        rowid_Accao = dataList[i][1].ToString(),
                        ref_Accao = dataList[i][2].ToString(),
                        codigo_Curso = dataList[i][3].ToString(),
                        nome_Curso = dataList[i][4].ToString(),
                        num_Sessao = dataList[i][5].ToString(),
                        data = DateTime.Parse(dataList[i][6].ToString()),
                        hora_Inicio = DateTime.Parse(dataList[i][7].ToString()),
                        hora_Fim = DateTime.Parse(dataList[i][8].ToString()),
                        rowid_Modulo = dataList[i][9].ToString(),
                        modulo = dataList[i][10].ToString(),
                        codigo_Formador = dataList[i][11].ToString(),
                        nome_Formador = dataList[i][12].ToString(),
                        versao_Data = dataList[i][13].ToString(),
                        local = dataList[i][14].ToString(),
                        tipo_Curso = dataList[i][15].ToString(),
                        estado = dataList[i][16].ToString(),
                        select = false,
                        codigo_Formador2 = dataList[i][18].ToString()
                    };
                    sessoesHT.Add(obj);
                }
            }
        }
        private void GetFeriadosHT()
        {
            feriadosHT.Clear();
            string subQuery = "select Data from TBGerFeriados";
            Connects.HTlocalConnect.ConnInit();
            SqlDataAdapter adapter = new SqlDataAdapter(subQuery, Connects.HTlocalConnect.Conn);
            DataTable subData = new DataTable(); adapter.Fill(subData);
            List<DataRow> dataList = subData.AsEnumerable().ToList();
            Connects.HTlocalConnect.ConnEnd();
            Connects.CloseAll();
            if (dataList.Count > 0)
            {
                for (int i = 0; i < dataList.Count; i++)
                {
                    var obj = new FeriadosHT()
                    {
                        Data = DateTime.Parse(dataList[i][0].ToString()),
                    };
                    feriadosHT.Add(obj);
                }
            }
        }
        private void Lista()
        {
            label1.Text = DB.selectCurso.Descricao + "_" + actualRef;
            avaliacoes.Clear();
            gridControl2.DataSource = null;
            double d = (double)modulosHT.Count / 2;
            int testeCount = 0, resto = (modulosHT.Count % 2);
            int modulosOrdem = 0, ordemgeral = 0, sessao = 0, metademodulos = (int)Math.Ceiling(d);
            var itemMetade = modulosHT.Where(x => x.Ordem.ToString() == metademodulos.ToString()).First();
            var itemUltimo = modulosHT.Where(x => x.Ordem.ToString() == modulosHT.Count.ToString()).First();

            DateTime? DataI1_EN = VerifFeriado(Get15diasApos(itemMetade.Data_Fim)),
                      DataF1_EN = VerifFeriado(DataI1_EN.Value.AddDays(1)),
                      DataI1_ER = VerifFeriado(DataF1_EN.Value.AddDays(6)),
                      Data1_ER = VerifFeriado(DataI1_ER.Value.AddDays(-5)),
                      DataF1_ER = VerifFeriado(DataI1_ER.Value.AddDays(1));

            DateTime? DataI2_EN = VerifFeriado(Get15diasApos(itemUltimo.Data_Fim)),
                      DataF2_EN = VerifFeriado(DataI2_EN.Value.AddDays(1)),
                      DataI2_ER = VerifFeriado(DataF2_EN.Value.AddDays(6)),
                      Data2_ER = VerifFeriado(DataI2_ER.Value.AddDays(-5)),
                      DataF2_ER = VerifFeriado(DataI2_ER.Value.AddDays(1));

            DateTime EEi1 = VerifFeriado(DataF1_ER.Value.AddDays(6)),
                     EEf1 = VerifFeriado(EEi1.AddDays(1));
            DateTime EEi2 = VerifFeriado(DataF2_ER.Value.AddDays(6)),
                     EEf2 = VerifFeriado(EEi2.AddDays(1));

            for (int i = 1; i <= metademodulos; i++)
            {
                testeCount++;
                modulosOrdem++;
                var item = modulosHT.Where(x => x.Ordem.ToString() == modulosOrdem.ToString()).First();

                string subQuery1 = "SELECT PARAM_AVAL, DESCRICAO, FACTOR_POND FROM TBForCursosParamsAvalQtva where nivel_modulo = " + item.Nivel + " AND Codigo_Curso = '" + actualCode + "' AND Momento_Aval = 1 ORDER BY versao_rowid";
                SqlDataAdapter adapter1 = new SqlDataAdapter(subQuery1, Connects.HTlocalConnect.Conn);
                DataTable subData1 = new DataTable();
                adapter1.Fill(subData1);
                List<DataRow> Metolologias1 = subData1.AsEnumerable().ToList();
                Connects.HTlocalConnect.ConnEnd();

                //Adiciona o módulo número impar no ultimo momento de avaliação 
                if (i == metademodulos && resto == 1)
                {
                    modulosOrdem++;
                    var item3 = modulosHT.Where(x => x.Ordem.ToString() == modulosHT.Count.ToString()).First();
                    string subQuery3 = "SELECT PARAM_AVAL, DESCRICAO, FACTOR_POND FROM TBForCursosParamsAvalQtva where nivel_modulo = " + item3.Nivel + " AND Codigo_Curso = '" + actualCode + "' AND Momento_Aval = 1 ORDER BY versao_rowid";
                    SqlDataAdapter adapter3 = new SqlDataAdapter(subQuery3, Connects.HTlocalConnect.Conn);
                    DataTable subData3 = new DataTable();
                    adapter3.Fill(subData3);
                    List<DataRow> Metolologias3 = subData3.AsEnumerable().ToList();
                    Connects.HTlocalConnect.ConnEnd();
                    avaliacoesativ.Clear();

                    if (Metolologias3.Count > 0)
                    {
                        for (int j = 0; j < Metolologias3.Count; j++)
                        {
                            var obj3 = new Avaliacao()
                            {
                                ID = 0,
                                Codigo_Curso = actualCode,
                                RefAcao = actualRef,
                                MomentoAv = "AVALIAÇÃO " + ToRoman(testeCount),
                                Modulo = item3.Descricao,
                                ModuloID = item3.Modulo_ID,
                                Ordem = ordemgeral++,
                                PesoCF = item3.Peso_Aval,
                                AvaliacaoN = testeCount,
                                select = false
                            };

                            obj3.PesoAv = (float.Parse(Metolologias3[j][2].ToString()) * 100) + "%";
                            obj3.Metodologia = Metolologias3[j][1].ToString();

                            string param = Metolologias3[j][0].ToString(), descricao = Metolologias3[j][1].ToString();

                            if (param == "EN [Teste]")
                            {
                                obj3.DataI_EN = VerifFeriado(Get15diasApos(item3.Data_Fim));
                                obj3.DataF_EN = VerifFeriado(obj3.DataI_EN.Value.AddDays(1));
                                obj3.DataI_ER = VerifFeriado(obj3.DataF_EN.Value.AddDays(6));
                                obj3.Data_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(-5));
                                obj3.DataF_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(1));
                                if (i <= metademodulos / 2)
                                {
                                    obj3.DataI_EE = EEi1;
                                    obj3.DataF_EE = EEf1;
                                    obj3.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                }
                                else
                                {
                                    obj3.DataI_EE = EEi2;
                                    obj3.DataF_EE = EEf2;
                                    obj3.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                }
                                obj3.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                obj3.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                obj3.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                            }
                            else if (param == "EN [Atvd + Vd]")
                            {
                                if (sessao == 0)
                                {
                                     //Data inicio mod
                                    string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item3.Descricao + "' order by Hora_Inicio asc";
                                    SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                    DataTable subData23 = new DataTable();
                                    adapterSessao23.Fill(subData23);
                                    List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                    Connects.HTlocalConnect.ConnEnd();
                                    if (Sessao23.Count > 0)
                                    {
                                        obj3.DataF_EN = ((DateTime)Sessao23[0][0]).AddDays(-1);
                                    }
                                    obj3.DataI_EN = item3.Data_Inicio;
                                }
                                else
                                {
                                    int contasessoes = modulosHTSessoes.Where(x => x.Ordem.ToString() == item3.Ordem.ToString()).Count();
                                    if (contasessoes > 1)
                                    {
                                        int atividadeordem = 0;
                                        foreach (var sessaoid in modulosHTSessoes.Where(x => x.Ordem.ToString() == item3.Ordem.ToString()).OrderBy(x => x.Data_Inicio))
                                        {
                                            var item4 = modulosHTSessoes.Where(x => x.Teste_ID == (sessaoid.Teste_ID - 1)).First();
                                            if (atividadeordem == 0)
                                            {
                                                obj3.DataI_EN = item4.Data_Fim.AddDays(+1);
                                                obj3.DataF_EN = sessaoid.Data_Fim.AddDays(-1);
                                                obj3.Metodologia = Metolologias3[j][1].ToString() + " " + ToRoman(atividadeordem + 1);
                                                obj3.Atividade = 1;
                                            }
                                            else
                                            {
                                                var objsessao2 = new Avaliacao()
                                                {
                                                    ID = 0,
                                                    Codigo_Curso = actualCode,
                                                    RefAcao = actualRef,
                                                    MomentoAv = "AVALIAÇÃO " + ToRoman(testeCount),
                                                    Modulo = item3.Descricao,
                                                    ModuloID = item3.Modulo_ID,
                                                    Ordem = ordemgeral++,
                                                    PesoCF = item3.Peso_Aval,
                                                    AvaliacaoN = testeCount,
                                                    select = false,
                                                    PesoAv = (float.Parse(Metolologias3[j][2].ToString()) * 100) + "%",
                                                    Metodologia = Metolologias3[j][1].ToString() + " " + ToRoman(atividadeordem + 1),
                                                    DataI_EN = item4.Data_Fim.AddDays(+1),
                                                    DataF_EN = sessaoid.Data_Fim.AddDays(-1),
                                                    Atividade = 1
                                                };
                                                avaliacoesativ.Add(objsessao2);
                                            }
                                            atividadeordem++;
                                        }
                                    }
                                    else
                                    {
                                        var item4 = modulosHT.Where(x => x.Ordem.ToString() == sessao.ToString()).First();
                                        obj3.DataI_EN = item4.Data_Fim.AddDays(+1);
                                        obj3.DataF_EN = item3.Data_Fim.AddDays(-1);
                                    }
                                }
                                sessao++;
                            }
                            else if (param == "EN [Atvd]")
                            {
                                //Data inicio mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item3.Descricao + "' order by Hora_Inicio asc";
                                SqlDataAdapter adapterSessao2 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData2 = new DataTable();
                                adapterSessao2.Fill(subData2);
                                List<DataRow> Sessao2 = subData2.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao2.Count > 0)
                                {
                                    obj3.DataI_EN = VerifFeriado((DateTime)Sessao2[0][0]);
                                    obj3.DataF_EN = VerifFeriado(VerifFeriado(Get15diasApos(item3.Data_Fim)).AddDays(1));
                                    obj3.DataI_ER = VerifFeriado(obj3.DataF_EN.Value.AddDays(6));
                                    obj3.Data_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(-5));
                                    obj3.DataF_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj3.DataI_EE = EEi1;
                                        obj3.DataF_EE = EEf1;
                                        obj3.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj3.DataI_EE = EEi2;
                                        obj3.DataF_EE = EEf2;
                                        obj3.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj3.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj3.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj3.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (param == "EN [Proj]")
                            {
                                obj3.DataI_EN = VerifFeriado(item3.Data_Fim.AddDays(1));
                                obj3.DataF_EN = VerifFeriado(obj3.DataI_EN.Value.AddMonths(1));
                            }
                            else if (VerifSP(param, descricao, DB.selectCurso.Codigo_Curso))
                            {
                                //Data inicio mod
                                string subQuerySessao2 = "SELECT s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item3.Descricao + "' order by Hora_Inicio asc";
                                SqlDataAdapter adapterSessao2 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData2 = new DataTable();
                                adapterSessao2.Fill(subData2);
                                List<DataRow> Sessao2 = subData2.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao2.Count > 0)
                                {
                                    obj3.DataI_EN = (DateTime)Sessao2[0][0];
                                    obj3.DataF_EN = (DateTime)Sessao2[Sessao2.Count-1][0];
                                    obj3.DataI_ER = VerifFeriado(obj3.DataF_EN.Value.AddDays(6));
                                    obj3.Data_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(-5));
                                    obj3.DataF_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj3.DataI_EE = EEi1;
                                        obj3.DataF_EE = EEf1;
                                        obj3.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj3.DataI_EE = EEi2;
                                        obj3.DataF_EE = EEf2;
                                        obj3.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj3.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj3.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj3.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (VerifEstCs(param, descricao, DB.selectCurso.Codigo_Curso))
                            {
                                //Data fim mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item3.Descricao + "' order by Hora_Inicio desc";
                                SqlDataAdapter adapterSessao2 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData2 = new DataTable();
                                adapterSessao2.Fill(subData2);
                                List<DataRow> Sessao2 = subData2.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao2.Count > 0)
                                {
                                    obj3.DataI_EN = VerifFeriado(((DateTime)Sessao2[0][0]).AddDays(1));
                                    obj3.DataF_EN = VerifFeriado(obj3.DataI_EN.Value.AddDays(7));
                                    obj3.DataI_ER = VerifFeriado(Get15diasApos(obj3.DataF_EN.Value));
                                    obj3.Data_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(-5));
                                    obj3.DataF_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj3.DataI_EE = EEi1;
                                        obj3.DataF_EE = EEf1;
                                        obj3.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj3.DataI_EE = EEi2;
                                        obj3.DataF_EE = EEf2;
                                        obj3.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj3.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj3.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj3.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (VerifRltTrabInd(param, descricao))
                            {
                                //Data fim mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item3.Descricao + "' order by Hora_Inicio desc";
                                SqlDataAdapter adapterSessao2 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData2 = new DataTable();
                                adapterSessao2.Fill(subData2);
                                List<DataRow> Sessao2 = subData2.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao2.Count > 0)
                                {
                                    obj3.DataI_EN = VerifFeriado(((DateTime)Sessao2[0][0]).AddDays(1));
                                    obj3.DataF_EN = VerifFeriado(obj3.DataI_EN.Value.AddDays(14));
                                    obj3.DataI_ER = VerifFeriado(Get15diasApos(obj3.DataF_EN.Value));
                                    obj3.Data_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(-5));
                                    obj3.DataF_ER = VerifFeriado(obj3.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj3.DataI_EE = EEi1;
                                        obj3.DataF_EE = EEf1;
                                        obj3.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj3.DataI_EE = EEi2;
                                        obj3.DataF_EE = EEf2;
                                        obj3.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj3.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj3.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj3.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            avaliacoes.Add(obj3);
                        }
                    }
                }
                else
                {
                    modulosOrdem++;
                    var item2 = modulosHT.Where(x => x.Ordem.ToString() == modulosOrdem.ToString()).First();
                    string subQuery2 = "SELECT PARAM_AVAL, DESCRICAO, FACTOR_POND FROM TBForCursosParamsAvalQtva where nivel_modulo = " + item2.Nivel + " AND Codigo_Curso = '" + actualCode + "' AND Momento_Aval = 1 ORDER BY versao_rowid";
                    SqlDataAdapter adapter2 = new SqlDataAdapter(subQuery2, Connects.HTlocalConnect.Conn);
                    DataTable subData2 = new DataTable();
                    adapter2.Fill(subData2);
                    List<DataRow> Metolologias2 = subData2.AsEnumerable().ToList();
                    Connects.HTlocalConnect.ConnEnd();

                    if (Metolologias1.Count > 0)
                    {
                        for (int j = 0; j < Metolologias1.Count; j++)
                        {
                            var obj = new Avaliacao()
                            {
                                ID = 0,
                                Codigo_Curso = actualCode,
                                RefAcao = actualRef,
                                MomentoAv = "AVALIAÇÃO " + ToRoman(testeCount),
                                Modulo = item.Descricao,
                                ModuloID = item.Modulo_ID,
                                Ordem = ordemgeral++,
                                PesoCF = item.Peso_Aval,
                                AvaliacaoN = testeCount,
                                select = false
                            };
                            
                            string param = Metolologias1[j][0].ToString(), descricao = Metolologias1[j][1].ToString();
                            obj.PesoAv = (float.Parse(Metolologias1[j][2].ToString()) * 100) + "%";
                            obj.Metodologia = Metolologias1[j][1].ToString();

                            if (param == "EN [Teste]")
                            {
                                obj.DataI_EN = VerifFeriado(Get15diasApos(item2.Data_Fim));
                                obj.DataF_EN = VerifFeriado(obj.DataI_EN.Value.AddDays(1));
                                obj.DataI_ER = VerifFeriado(obj.DataF_EN.Value.AddDays(6));
                                obj.Data_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(-5));
                                obj.DataF_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(1));
                                if (i <= metademodulos / 2)
                                {
                                    obj.DataI_EE = EEi1;
                                    obj.DataF_EE = EEf1;
                                    obj.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                }
                                else
                                {
                                    obj.DataI_EE = EEi2;
                                    obj.DataF_EE = EEf2;
                                    obj.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                }
                                obj.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                obj.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                obj.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                            }
                            else if (param == "EN [Atvd + Vd]")
                            {
                                if (sessao == 0)
                                {
                                    //Data inicio mod
                                    string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item.Descricao + "' order by Hora_Inicio asc";
                                    SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                    DataTable subData23 = new DataTable();
                                    adapterSessao23.Fill(subData23);
                                    List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                    Connects.HTlocalConnect.ConnEnd();
                                    if (Sessao23.Count > 0)
                                        obj.DataF_EN = ((DateTime)Sessao23[0][0]).AddDays(-1);

                                    obj.DataI_EN = item.Data_Inicio;
                                }
                                else
                                {
                                    int contasessoes = modulosHTSessoes.Where(x => x.Ordem.ToString() == item.Ordem.ToString()).Count();
                                    if (contasessoes > 1)
                                    {
                                        int atividadeordem = 0;
                                        foreach (var sessaoid in modulosHTSessoes.Where(x => x.Ordem.ToString() == item.Ordem.ToString()).OrderBy(x => x.Data_Inicio))
                                        {
                                            var item4 = modulosHTSessoes.Where(x => x.Teste_ID == (sessaoid.Teste_ID - 1)).First();
                                            if (atividadeordem == 0)
                                            {
                                                obj.DataI_EN = item4.Data_Fim.AddDays(+1);
                                                obj.DataF_EN = sessaoid.Data_Fim.AddDays(-1);
                                                obj.Metodologia = Metolologias1[j][1].ToString() + " " + ToRoman(atividadeordem + 1);
                                                obj.Atividade = 1;
                                            }
                                            else
                                            {
                                                var objsessao2 = new Avaliacao()
                                                {
                                                    ID = 0,
                                                    Codigo_Curso = actualCode,
                                                    RefAcao = actualRef,
                                                    MomentoAv = "AVALIAÇÃO " + ToRoman(testeCount),
                                                    Modulo = item.Descricao,
                                                    ModuloID = item.Modulo_ID,
                                                    Ordem = ordemgeral++,
                                                    PesoCF = item.Peso_Aval,
                                                    AvaliacaoN = testeCount,
                                                    select = false,
                                                    PesoAv = (float.Parse(Metolologias1[j][2].ToString()) * 100) + "%",
                                                    Metodologia = Metolologias1[j][1].ToString() + " " + ToRoman(atividadeordem + 1),
                                                    DataI_EN = item4.Data_Fim.AddDays(+1),
                                                    DataF_EN = sessaoid.Data_Fim.AddDays(-1),
                                                    Atividade = 1
                                                };
                                                avaliacoesativ.Add(objsessao2);
                                            }
                                            atividadeordem++;
                                        }
                                    }
                                    else
                                    {
                                        var item4 = modulosHT.Where(x => x.Ordem.ToString() == sessao.ToString()).First();
                                        obj.DataI_EN = item4.Data_Fim.AddDays(+1);
                                        obj.DataF_EN = item.Data_Fim.AddDays(-1);
                                    }
                                }
                                sessao++;
                            }
                            else if (param == "EN [Atvd]")
                            {
                                //Data inicio mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item.Descricao + "' order by Hora_Inicio asc";
                                SqlDataAdapter adapterSessao2 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                subData2 = new DataTable();
                                adapterSessao2.Fill(subData2);
                                List<DataRow> Sessao2 = subData2.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();

                                if (Sessao2.Count > 0)
                                {
                                    obj.DataI_EN = VerifFeriado((DateTime)Sessao2[0][0]);
                                    obj.DataF_EN = VerifFeriado(VerifFeriado(Get15diasApos(item2.Data_Fim)).AddDays(1));
                                    obj.DataI_ER = VerifFeriado(obj.DataF_EN.Value.AddDays(6));
                                    obj.Data_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(-5));
                                    obj.DataF_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj.DataI_EE = EEi1;
                                        obj.DataF_EE = EEf1;
                                        obj.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj.DataI_EE = EEi2;
                                        obj.DataF_EE = EEf2;
                                        obj.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (param == "EN [Proj]")
                            {
                                obj.DataI_EN = VerifFeriado(item2.Data_Fim.AddDays(1));
                                obj.DataF_EN = VerifFeriado(obj.DataI_EN.Value.AddMonths(1));
                            }
                            else if (VerifSP(param, descricao, DB.selectCurso.Codigo_Curso))
                            {
                                //Data inicio mod
                                string subQuerySessao2 = "SELECT s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item.Descricao + "' order by Hora_Inicio asc";
                                SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData23 = new DataTable();
                                adapterSessao23.Fill(subData23);
                                List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao23.Count > 0)
                                {
                                    obj.DataI_EN = (DateTime)Sessao23[0][0];
                                    obj.DataF_EN = (DateTime)Sessao23[Sessao23.Count-1][0];
                                    obj.DataI_ER = VerifFeriado(Get15diasApos(obj.DataF_EN.Value));
                                    obj.Data_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(-5));
                                    obj.DataF_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj.DataI_EE = EEi1;
                                        obj.DataF_EE = EEf1;
                                        obj.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj.DataI_EE = EEi2;
                                        obj.DataF_EE = EEf2;
                                        obj.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (VerifEstCs(param, descricao, DB.selectCurso.Codigo_Curso))
                            {
                                //Data inicio mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item.Descricao + "' order by Hora_Inicio desc";
                                SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData23 = new DataTable();
                                adapterSessao23.Fill(subData23);
                                List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao23.Count > 0)
                                {
                                    obj.DataI_EN = VerifFeriado(((DateTime)Sessao23[0][0]).AddDays(1));
                                    obj.DataF_EN = VerifFeriado(obj.DataI_EN.Value.AddDays(7));
                                    obj.DataI_ER = VerifFeriado(Get15diasApos(obj.DataF_EN.Value));
                                    obj.Data_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(-5));
                                    obj.DataF_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj.DataI_EE = EEi1;
                                        obj.DataF_EE = EEf1;
                                        obj.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj.DataI_EE = EEi2;
                                        obj.DataF_EE = EEf2;
                                        obj.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (VerifRltTrabInd(param, descricao))
                            {
                                //Data fim mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item.Descricao + "' order by Hora_Inicio desc";
                                SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData23 = new DataTable();
                                adapterSessao23.Fill(subData23);
                                List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao23.Count > 0)
                                {
                                    obj.DataI_EN = VerifFeriado(((DateTime)Sessao23[0][0]).AddDays(1));
                                    obj.DataF_EN = VerifFeriado(obj.DataI_EN.Value.AddDays(14));
                                    obj.DataI_ER = VerifFeriado(Get15diasApos(obj.DataF_EN.Value));
                                    obj.Data_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(-5));
                                    obj.DataF_ER = VerifFeriado(obj.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj.DataI_EE = EEi1;
                                        obj.DataF_EE = EEf1;
                                        obj.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj.DataI_EE = EEi2;
                                        obj.DataF_EE = EEf2;
                                        obj.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            avaliacoes.Add(obj);
                        }
                    }
                    if (Metolologias2.Count > 0)
                    {
                        for (int j = 0; j < Metolologias2.Count; j++)
                        {
                            var obj2 = new Avaliacao()
                            {
                                ID = 0,
                                Codigo_Curso = actualCode,
                                RefAcao = actualRef,
                                MomentoAv = "AVALIAÇÃO " + ToRoman(testeCount),
                                Modulo = item2.Descricao,
                                ModuloID = item2.Modulo_ID,
                                Ordem = ordemgeral++,
                                PesoCF = item2.Peso_Aval,
                                AvaliacaoN = testeCount,
                                select = false
                            };
                            string param = Metolologias2[j][0].ToString(), descricao = Metolologias2[j][1].ToString();
                            obj2.PesoAv = (float.Parse(Metolologias2[j][2].ToString()) * 100) + "%";
                            obj2.Metodologia = Metolologias2[j][1].ToString();

                            if (param == "EN [Teste]")
                            {
                                obj2.DataI_EN = VerifFeriado(Get15diasApos(item2.Data_Fim));
                                obj2.DataF_EN = VerifFeriado(obj2.DataI_EN.Value.AddDays(1));
                                obj2.DataI_ER = VerifFeriado(obj2.DataF_EN.Value.AddDays(6));
                                obj2.Data_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(-5));
                                obj2.DataF_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(1));
                                if (i <= metademodulos / 2)
                                {
                                    obj2.DataI_EE = EEi1;
                                    obj2.DataF_EE = EEf1;
                                    obj2.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                }
                                else
                                {
                                    obj2.DataI_EE = EEi2;
                                    obj2.DataF_EE = EEf2;
                                    obj2.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                }
                                obj2.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                obj2.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                obj2.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                            }
                            else if (param == "EN [Atvd + Vd]")
                            {
                                if (sessao == 0)
                                {
                                     //Data inicio mod
                                    string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item2.Descricao + "' order by Hora_Inicio asc";
                                    SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                    DataTable subData23 = new DataTable();
                                    adapterSessao23.Fill(subData23);
                                    List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                    Connects.HTlocalConnect.ConnEnd();
                                    if (Sessao23.Count > 0)
                                    {
                                        obj2.DataF_EN = ((DateTime)Sessao23[0][0]).AddDays(-1);
                                    }
                                    obj2.DataI_EN = item2.Data_Inicio;
                                }
                                else
                                {
                                    int contasessoes = modulosHTSessoes.Where(x => x.Ordem.ToString() == item2.Ordem.ToString()).Count();
                                    if (contasessoes > 1)
                                    {
                                        int atividadeordem = 0;
                                        foreach (var sessaoid in modulosHTSessoes.Where(x => x.Ordem.ToString() == item2.Ordem.ToString()).OrderBy(x => x.Data_Inicio))
                                        {
                                            var item4 = modulosHTSessoes.Where(x => x.Teste_ID == (sessaoid.Teste_ID - 1)).First();
                                            if (atividadeordem == 0)
                                            {
                                                obj2.DataI_EN = item4.Data_Fim.AddDays(+1);
                                                obj2.DataF_EN = sessaoid.Data_Fim.AddDays(-1);
                                                obj2.Metodologia = Metolologias2[j][1].ToString() + " " + ToRoman(atividadeordem + 1);
                                                obj2.Atividade = 1;
                                            }
                                            else
                                            {
                                                var objsessao2 = new Avaliacao()
                                                {
                                                    ID = 0,
                                                    Codigo_Curso = actualCode,
                                                    RefAcao = actualRef,
                                                    MomentoAv = "AVALIAÇÃO " + ToRoman(testeCount),
                                                    Modulo = item2.Descricao,
                                                    ModuloID = item2.Modulo_ID,
                                                    Ordem = ordemgeral++,
                                                    PesoCF = item2.Peso_Aval,
                                                    AvaliacaoN = testeCount,
                                                    select = false,
                                                    PesoAv = (float.Parse(Metolologias2[j][2].ToString()) * 100) + "%",
                                                    Metodologia = Metolologias2[j][1].ToString() + " " + ToRoman(atividadeordem + 1),
                                                    DataI_EN = item4.Data_Fim.AddDays(+1),
                                                    DataF_EN = sessaoid.Data_Fim.AddDays(-1),
                                                    Atividade = 1
                                                };
                                                avaliacoesativ.Add(objsessao2);
                                            }
                                            atividadeordem++;
                                        }
                                    }
                                    else
                                    {
                                        var item4 = modulosHT.Where(x => x.Ordem.ToString() == sessao.ToString()).First();
                                        obj2.DataI_EN = item4.Data_Fim.AddDays(+1);
                                        obj2.DataF_EN = item2.Data_Fim.AddDays(-1);
                                    }
                                }
                                sessao++;
                            }
                            else if (param == "EN [Atvd]")
                            {
                                //Data inicio mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item2.Descricao + "' order by Hora_Inicio asc";
                                SqlDataAdapter adapterSessao2 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                subData2 = new DataTable();
                                adapterSessao2.Fill(subData2);
                                List<DataRow> Sessao2 = subData2.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao2.Count > 0)
                                {
                                    obj2.DataI_EN = VerifFeriado((DateTime)Sessao2[0][0]);
                                    obj2.DataF_EN = VerifFeriado(VerifFeriado(Get15diasApos(item2.Data_Fim)).AddDays(1));
                                    obj2.DataI_ER = VerifFeriado(obj2.DataF_EN.Value.AddDays(6));
                                    obj2.Data_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(-5));
                                    obj2.DataF_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos)
                                    {
                                        obj2.DataI_EE = EEi1;
                                        obj2.DataF_EE = EEf1;
                                        obj2.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj2.DataI_EE = EEi2;
                                        obj2.DataF_EE = EEf2;
                                        obj2.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj2.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj2.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj2.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (param == "EN [Proj]")
                            {
                                obj2.DataI_EN = VerifFeriado(item2.Data_Fim.AddDays(1));
                                obj2.DataF_EN = VerifFeriado(obj2.DataI_EN.Value.AddMonths(1));
                            }
                            else if (VerifSP(param, descricao, DB.selectCurso.Codigo_Curso))
                            {
                                //Data inicio mod
                                string subQuerySessao2 = "SELECT s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item2.Descricao + "' order by Hora_Inicio asc";
                                SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData23 = new DataTable();
                                adapterSessao23.Fill(subData23);
                                List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao23.Count > 0)
                                {
                                    obj2.DataI_EN = (DateTime)Sessao23[0][0];
                                    obj2.DataF_EN = (DateTime)Sessao23[Sessao23.Count-1][0];
                                    obj2.DataI_ER = VerifFeriado(Get15diasApos(obj2.DataF_EN.Value));
                                    obj2.Data_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(-5));
                                    obj2.DataF_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj2.DataI_EE = EEi1;
                                        obj2.DataF_EE = EEf1;
                                        obj2.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj2.DataI_EE = EEi2;
                                        obj2.DataF_EE = EEf2;
                                        obj2.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj2.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj2.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj2.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (VerifEstCs(param, descricao, DB.selectCurso.Codigo_Curso))
                            {
                                //Data inicio mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item2.Descricao + "' order by Hora_Inicio desc";
                                SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData23 = new DataTable();
                                adapterSessao23.Fill(subData23);
                                List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao23.Count > 0)
                                {
                                    obj2.DataI_EN = VerifFeriado(((DateTime)Sessao23[0][0]).AddDays(1));
                                    obj2.DataF_EN = VerifFeriado(obj2.DataI_EN.Value.AddDays(7));
                                    obj2.DataI_ER = VerifFeriado(Get15diasApos(obj2.DataF_EN.Value));
                                    obj2.Data_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(-5));
                                    obj2.DataF_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj2.DataI_EE = EEi1;
                                        obj2.DataF_EE = EEf1;
                                        obj2.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj2.DataI_EE = EEi2;
                                        obj2.DataF_EE = EEf2;
                                        obj2.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj2.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj2.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj2.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            else if (VerifRltTrabInd(param, descricao))
                            {
                                //Data fim mod
                                string subQuerySessao2 = "SELECT top 1 s.Data FROM TBForSessoes s JOIN TBForAccoes a on a.versao_rowid = s.Rowid_Accao JOIN TBForModulos m on m.versao_rowid = s.rowid_modulo inner join TBForFormadores b on b.Codigo_Formador = s.Codigo_Formador WHERE a.ref_accao = '" + DB.selectCurso.Ref_Accao + "' and extra = 0 and Num_Sessao > 0 and Descricao = '" + item2.Descricao + "' order by Hora_Inicio desc";
                                SqlDataAdapter adapterSessao23 = new SqlDataAdapter(subQuerySessao2, Connects.HTlocalConnect.Conn);
                                DataTable subData23 = new DataTable();
                                adapterSessao23.Fill(subData23);
                                List<DataRow> Sessao23 = subData23.AsEnumerable().ToList();
                                Connects.HTlocalConnect.ConnEnd();
                                if (Sessao23.Count > 0)
                                {
                                    obj2.DataI_EN = VerifFeriado(((DateTime)Sessao23[0][0]).AddDays(1));
                                    obj2.DataF_EN = VerifFeriado(obj2.DataI_EN.Value.AddDays(14));
                                    obj2.DataI_ER = VerifFeriado(Get15diasApos(obj2.DataF_EN.Value));
                                    obj2.Data_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(-5));
                                    obj2.DataF_ER = VerifFeriado(obj2.DataI_ER.Value.AddDays(1));
                                    if (i <= metademodulos / 2)
                                    {
                                        obj2.DataI_EE = EEi1;
                                        obj2.DataF_EE = EEf1;
                                        obj2.Data_EE = VerifFeriado(EEi1.AddDays(-5));
                                    }
                                    else
                                    {
                                        obj2.DataI_EE = EEi2;
                                        obj2.DataF_EE = EEf2;
                                        obj2.Data_EE = VerifFeriado(EEi2.AddDays(-5));
                                    }
                                    obj2.DataI_EXT = VerifFeriado(EEf2.AddDays(6));
                                    obj2.Data_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(-5));
                                    obj2.DataF_EXT = VerifFeriado(VerifFeriado(EEf2.AddDays(6)).AddDays(1));
                                }
                            }
                            avaliacoes.Add(obj2);
                        }
                    }
                }
            }

            avaliacaocalculo.Clear();
            DataTable table = new DataTable();
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Codigo_Curso");
            table.Columns.Add("RefAcao");
            table.Columns.Add("MomentoAv");
            table.Columns.Add("Modulo");
            table.Columns.Add("ModuloID");
            table.Columns.Add("Ordem");
            table.Columns.Add("PesoCF");
            table.Columns.Add("Metodologia");
            table.Columns.Add("PesoAv");
            table.Columns.Add("DataI_EN");
            table.Columns.Add("DataF_EN");
            table.Columns.Add("DataI_ER");
            table.Columns.Add("DataF_ER");
            table.Columns.Add("DataI_EE");
            table.Columns.Add("DataF_EE");
            table.Columns.Add("DataI_EXT");
            table.Columns.Add("DataF_EXT");
            table.Columns.Add("Data_ER");
            table.Columns.Add("Data_EE");
            table.Columns.Add("Data_EXT");
            table.Columns.Add("AvaliacaoN");

            for (int j = 0; j < avaliacoes.Count; j++)
            {
                if (avaliacoes[j].Metodologia == "Avaliação Contínua")
                {
                    table.Rows.Add(new object[]
                    {
                        0,
                        actualCode,
                        actualRef,
                        avaliacoes[j].MomentoAv,
                        avaliacoes[j].Modulo,
                        avaliacoes[j].ModuloID,
                        avaliacoes[j].Ordem,
                        avaliacoes[j].PesoCF,
                        avaliacoes[j].Metodologia,
                        avaliacoes[j].PesoAv,
                        null, null, null, null, null, null, null, null, null, null, null,
                        avaliacoes[j].AvaliacaoN
                    });
                }
                else
                {
                    table.Rows.Add(new object[]
                    {
                        0,
                        actualCode,
                        actualRef,
                        avaliacoes[j].MomentoAv,
                        avaliacoes[j].Modulo,
                        avaliacoes[j].ModuloID,
                        avaliacoes[j].Ordem,
                        avaliacoes[j].PesoCF,
                        avaliacoes[j].Metodologia,
                        avaliacoes[j].PesoAv,
                        avaliacoes[j].DataI_EN.ToString() == "" ? "" : avaliacoes[j].DataI_EN.Value.ToShortDateString(),
                        avaliacoes[j].DataF_EN.ToString() == "" ? "" : avaliacoes[j].DataF_EN.Value.ToShortDateString(),
                        avaliacoes[j].DataI_ER.ToString() == "" ? "" : avaliacoes[j].DataI_ER.Value.ToShortDateString(),
                        avaliacoes[j].DataF_ER.ToString() == "" ? "" : avaliacoes[j].DataF_ER.Value.ToShortDateString(),
                        avaliacoes[j].DataI_EE.ToString() == "" ? "" : avaliacoes[j].DataI_EE.Value.ToShortDateString(),
                        avaliacoes[j].DataF_EE.ToString() == "" ? "" : avaliacoes[j].DataF_EE.Value.ToShortDateString(),
                        avaliacoes[j].DataI_EXT.ToString() == "" ? "" : avaliacoes[j].DataI_EXT.Value.ToShortDateString(),
                        avaliacoes[j].DataF_EXT.ToString() == "" ? "" : avaliacoes[j].DataF_EXT.Value.ToShortDateString(),
                        avaliacoes[j].Data_ER.ToString() == "" ? "" : avaliacoes[j].Data_ER.Value.ToShortDateString(),
                        avaliacoes[j].Data_EE.ToString() == "" ? "" : avaliacoes[j].Data_EE.Value.ToShortDateString(),
                        avaliacoes[j].Data_EXT.ToString() == "" ? "" : avaliacoes[j].Data_EXT.Value.ToShortDateString(),
                        avaliacoes[j].AvaliacaoN
                    });
                }
            }

            GridView view = gridControl2.MainView as GridView;
            gridControl2.DataSource = table;
            view.OptionsBehavior.Editable = false;
        }
        private bool VerifRltTrabInd(string param, string descricao)
        {
            if(param == "EN [Rlt]" || param == "EN [TrabInd]" ||                                 (param == "EN [Trab]" && (descricao == "Trabalho" || descricao == "Trabalho prático" || descricao == "Trabalho prático individual")))
                return true;
            return false;
        }
        private bool VerifEstCs(string param, string descricao, string codcurso)
        {
            if(param == "EN [EstCs]" || (param == "EN [Caso]" && descricao == "Estudo de caso" && codcurso.ToUpper() == "CTAV_V2"))
                return true;
            return false;
        }
        private bool VerifSP(string param, string descricao, string codcurso)
        {
            if(param == "EN [Sp]" || param == "EN [Simul]" || 
                (param == "EN [Caso]" && descricao == "Análise e Resolução de Casos Práticos") ||
                (param == "EN [Trab]" && descricao == "Plano de intervenção") ||
                (param == "EN [Caso]" && descricao == "Estudo de caso" && codcurso.ToUpper() == "PTTG_V3") ||
                (param == "EN [Trab]" && descricao == "Plano de intervenção" && codcurso.ToUpper() == "PTTG_V4") ||
                (param == "EN [Proj]" && descricao == "Projeto final" && codcurso.ToUpper() == "PPCLF_V5"))
                return true;
            return false;
        }
        public DateTime Get15diasApos(DateTime d)
        {
            DateTime da = DateTime.MinValue;
            if (d.DayOfWeek == DayOfWeek.Monday) { da = d.AddDays(13); };
            if (d.DayOfWeek == DayOfWeek.Tuesday) { da = d.AddDays(12); };
            if (d.DayOfWeek == DayOfWeek.Wednesday) { da = d.AddDays(11); };
            if (d.DayOfWeek == DayOfWeek.Thursday) { da = d.AddDays(10); };
            if (d.DayOfWeek == DayOfWeek.Friday) { da = d.AddDays(9); };
            if (d.DayOfWeek == DayOfWeek.Saturday) { da = d.AddDays(8); };
            if (d.DayOfWeek == DayOfWeek.Sunday) { da = d.AddDays(7); };
            return da;
        }
        public DateTime Get5diasApos(DateTime d)
        {
            DateTime da = DateTime.MinValue;
            if (d.DayOfWeek == DayOfWeek.Monday) { da = d.AddDays(7); };
            if (d.DayOfWeek == DayOfWeek.Tuesday) { da = d.AddDays(6); };
            if (d.DayOfWeek == DayOfWeek.Wednesday) { da = d.AddDays(5); };
            if (d.DayOfWeek == DayOfWeek.Thursday) { da = d.AddDays(11); };
            if (d.DayOfWeek == DayOfWeek.Friday) { da = d.AddDays(10); };
            if (d.DayOfWeek == DayOfWeek.Saturday) { da = d.AddDays(9); };
            if (d.DayOfWeek == DayOfWeek.Sunday) { da = d.AddDays(8); };
            return da;
        }
        public DateTime VerifFeriado(DateTime d)
        {
            GetFeriadosHT();
            int existe = feriadosHT.Where(x => x.Data.Date == d.Date).Count();
            if (existe > 0) { d = d.AddDays(1); }
            return d;
        }
        public void CalculaEspeciais()
        {
            DateTime testeLastDate = new DateTime(), datarecurso = new DateTime();
            if (avaliacaocalculo.Count > 0)
            {
                int metade = (avaliacaocalculo.Count / 2), resto = (avaliacaocalculo.Count % 2);
                metade += resto;

                testeLastDate = avaliacaocalculo[metade - 1].DataI_ER;
                datarecurso = testeLastDate.AddDays(7);

                while (datarecurso.DayOfWeek != DayOfWeek.Sunday)
                    datarecurso = datarecurso.AddDays(1);

                especial_1 = datarecurso;
                //Verifica se o dia selecionado é feriado
                int existe = feriadosHT.Where(x => x.Data.Date == especial_1.Date).Count();
                if (existe > 0)
                {
                    especial_1 = especial_1.AddDays(1);
                }

                especial_2 = avaliacaocalculo.OrderByDescending(x => x.DataI_ER).First().DataI_ER;
                especial_2 = especial_2.AddDays(7);

                while (especial_2.DayOfWeek != DayOfWeek.Sunday)
                    especial_2 = datarecurso.AddDays(1);
                //Verifica se o dia selecionado é feriado
                int existe2 = feriadosHT.Where(x => x.Data.Date == especial_2.Date).Count();
                if (existe2 > 0)
                {
                    especial_2 = especial_2.AddDays(1);
                }

                extraordinaria = especial_2;
                extraordinaria = extraordinaria.AddDays(8);
                while (extraordinaria.DayOfWeek != DayOfWeek.Sunday)
                    extraordinaria = extraordinaria.AddDays(1);
                //Verifica se o dia selecionado é feriado
                int existe3 = feriadosHT.Where(x => x.Data.Date == extraordinaria.Date).Count();
                if (existe3 > 0)
                {
                    extraordinaria = extraordinaria.AddDays(1);
                }

                DateTime datainsc_ER = new DateTime(), datainsc_EE = new DateTime();

                foreach (var item in avaliacoes)
                {
                    if (item.DataI_EN.ToString() != "")
                    {
                        if (item.AvaliacaoN <= metade)
                        {
                            item.DataI_EE = especial_1;
                            item.DataF_EE = especial_1.AddDays(1);
                        }
                        else
                        {
                            item.DataI_EE = especial_2;
                            item.DataF_EE = especial_2.AddDays(1);
                        }
                    }
                }

                foreach (var item in avaliacoes)
                {
                    if (item.DataI_ER.ToString() != "" && item.DataI_EE.ToString() != "")
                    {
                        //Verifica data de inscrição
                        datainsc_ER = item.DataI_ER.Value.AddDays(-4);
                        datainsc_EE = item.DataI_EE.Value.AddDays(-4);
                        extraordinariaINS = extraordinaria.AddDays(-4);

                        //Verifica se o dia selecionado é feriado
                        int existe5 = feriadosHT.Where(x => x.Data.Date == datainsc_ER.Date).Count();
                        if (existe5 > 0)
                        {
                            datainsc_ER = datainsc_ER.AddDays(-1);
                        }
                        int existe6 = feriadosHT.Where(x => x.Data.Date == datainsc_EE.Date).Count();
                        if (existe6 > 0)
                        {
                            datainsc_EE = datainsc_EE.AddDays(-1);
                        }
                        int existe7 = feriadosHT.Where(x => x.Data.Date == extraordinariaINS.Date).Count();
                        if (existe7 > 0)
                        {
                            extraordinariaINS = extraordinariaINS.AddDays(-1);
                        }

                        item.Data_ER = datainsc_ER;
                        item.Data_EE = datainsc_EE;
                    }
                }
            }
        }
        public string ToRoman(int number)
        {
            if (number < 0 || number > 3999) return string.Empty;
            if (number < 1) return string.Empty;
            if (number >= 1000) return "M" + ToRoman(number - 1000);
            if (number >= 900) return "CM" + ToRoman(number - 900);
            if (number >= 500) return "D" + ToRoman(number - 500);
            if (number >= 400) return "CD" + ToRoman(number - 400);
            if (number >= 100) return "C" + ToRoman(number - 100);
            if (number >= 90) return "XC" + ToRoman(number - 90);
            if (number >= 50) return "L" + ToRoman(number - 50);
            if (number >= 40) return "XL" + ToRoman(number - 40);
            if (number >= 10) return "X" + ToRoman(number - 10);
            if (number >= 9) return "IX" + ToRoman(number - 9);
            if (number >= 5) return "V" + ToRoman(number - 5);
            if (number >= 4) return "IV" + ToRoman(number - 4);
            if (number >= 1) return "I" + ToRoman(number - 1);
            return string.Empty;
        }
        private static int ConvertRomanToInt(String romanNumeral)
        {
            Dictionary<Char, Int32> romanMap = new Dictionary<char, int>
            {
                {'I', 1},
                {'V', 5},
                {'X', 10},
                {'L', 50},
                {'C', 100},
                {'D', 500},
                {'M', 1000}
            };
            Int32 result = 0;
            for (Int32 index = romanNumeral.Length - 1, last = 0; index >= 0; index--)
            {
                Int32 current = romanMap[romanNumeral[index]];
                result += (current < last ? -current : current);
                last = current;
            }
            return result;
        }
        private void Modulos_SelectedIndexChanged(object sender, EventArgs e)
        {
            ModuloID.Text = modulosHT.Where(x => x.Descricao == modulos.Text).First().Modulo_ID.ToString();
            pesocf.Text = modulosHT.Where(x => x.Descricao == modulos.Text).First().Peso_Aval.ToString();
        }
        private void SimpleButton3_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            bool faz = false;

            if (avaliacao.Text == "") return;
            if (avaliacao.Text != "")
            {
                if ((ConvertRomanToInt(avaliacao.Text)) == 0) return;
            }

            GridView view = gridControl2.MainView as GridView;
            for (int i = 0; i < view.DataRowCount; i++)
            {
                DataRow row = view.GetDataRow(i);
                if (row.ItemArray[0].ToString() == "0")
                {
                    MessageBox.Show("Clique em GUARDAR para continuar.", (string)Properties.Settings.Default["title"]);
                    return;
                }
            }

            string subQuery = "SELECT * FROM CalendarioAV WHERE METODOLOGIA = '" + metodologia.Text + "' AND MODULOID = " + ModuloID.Text + " AND AVALIACAON = " + ConvertRomanToInt(avaliacao.Text) + " AND Codigo_Curso = '" + actualCode + "' AND REFACAO = '" + actualRef + "' ORDER BY ORDEM";
            if (ID.Text != "")
                subQuery = "SELECT * FROM CalendarioAV WHERE ID <> " + ID.Text + " AND METODOLOGIA = '" + metodologia.Text + "' AND MODULOID = " + ModuloID.Text + " AND AVALIACAON = " + ConvertRomanToInt(avaliacao.Text) + " AND Codigo_Curso = '" + actualCode + "' AND REFACAO = '" + actualRef + "' ORDER BY ORDEM";

            SqlDataAdapter adapter = new SqlDataAdapter(subQuery, Connects.localConnect.Conn);
            DataTable subData = new DataTable();
            adapter.Fill(subData);
            List<DataRow> avaliacoesguardadas = subData.AsEnumerable().ToList();
            Connects.HTlocalConnect.ConnEnd();

            if (avaliacoesguardadas.Count > 0)
            {
                MessageBox.Show("O Momento de Avaliação já existe.", (string)Properties.Settings.Default["title"]);
                return;
            }

            if (ID.Text != "")
            {
                if (ModuloID.Text == "") return;
                if (metodologia.Text == "") return;
                if (PesoAv.Text == "" && ModuloID.Text != "0")
                { 
                    MessageBox.Show("Deve indicar o peso da avaliação.", (string)Properties.Settings.Default["title"]); 
                    return; 
                }
                if (PesoAv.Text == "" && ModuloID.Text == "0") PesoAv.Text = "0";
                if (pesocf.Text == "" && ModuloID.Text == "0") pesocf.Text = "0";

                Connects.localConnect.ConnInit();
                string format = "yyyy-MM-dd";
                subQuery = "UPDATE CalendarioAV SET Codigo_Curso = '" + actualCode + "', RefAcao = '" + actualRef + "', MomentoAv = '" + "AVALIAÇÃO " + avaliacao.Text + "', ModuloID = " + ModuloID.Text + ", PesoCF = " + pesocf.Text.Replace(",", ".") + ", Metodologia = '" + metodologia.Text + "', ";
                bool res = int.TryParse(PesoAv.Text, out int a);

                if (res) subQuery += "PesoAv = '" + PesoAv.Text + "%" + "', ";
                else subQuery += "PesoAv = '" + PesoAv.Text + "', ";

                subQuery += "AvaliacaoN = " + ConvertRomanToInt(avaliacao.Text) + ", ";

                if (op_datai_en.Checked == false) subQuery += "DataI_EN = Null,";
                else subQuery += "DataI_EN = '" + DataI_EN.Value.ToString(format) + "',";
                if (op_dataf_en.Checked == false) subQuery += "DataF_EN = Null,";
                else subQuery += "DataF_EN = '" + DataF_EN.Value.ToString(format) + "',";
                if (op_datai_er.Checked == false) subQuery += "DataI_ER = Null,";
                else subQuery += "DataI_ER = '" + DataI_ER.Value.ToString(format) + "',";
                if (op_dataf_er.Checked == false) subQuery += "DataF_ER = Null,";
                else subQuery += "DataF_ER = '" + DataF_ER.Value.ToString(format) + "',";
                if (op_datai_ee.Checked == false) subQuery += "DataI_EE = Null,";
                else subQuery += "DataI_EE = '" + DataI_EE.Value.ToString(format) + "',";
                if (op_dataf_ee.Checked == false) subQuery += "DataF_EE = Null,";
                else subQuery += "DataF_EE = '" + DataF_EE.Value.ToString(format) + "',";
                if (op_DataI_EXT.Checked == false) subQuery += "DataI_EXT = Null,";
                else subQuery += "DataI_EXT = '" + DataI_EXT.Value.ToString(format) + "',";
                if (op_DataF_EXT.Checked == false) subQuery += "DataF_EXT = Null,";
                else subQuery += "DataF_EXT = '" + DataF_EXT.Value.ToString(format) + "',";
                if (op_data_er.Checked == false) subQuery += "data_er = Null,";
                else subQuery += "data_er = '" + data_er.Value.ToString(format) + "',";
                if (op_data_ee.Checked == false) subQuery += "data_ee = Null,";
                else subQuery += "data_ee = '" + data_ee.Value.ToString(format) + "',";
                if (op_data_ext.Checked == false) subQuery += "data_ext = Null ";
                else subQuery += "data_ext = '" + data_ext.Value.ToString(format) + "' ";

                subQuery += " WHERE ID = " + ID.Text;
                SqlCommand insertModulo = new SqlCommand(subQuery, Connects.localConnect.Conn);
                insertModulo.CommandText = subQuery;
                insertModulo.ExecuteNonQuery();
                Connects.localConnect.ConnEnd();
                faz = true;
            }
            else
            {
                if (avaliacao.Text == "") return;
                if (ModuloID.Text == "") return;
                if (metodologia.Text == "") return;
                if (PesoAv.Text == "" && ModuloID.Text != "0")
                { 
                    MessageBox.Show("Deve indicar o peso da avaliação.", (string)Properties.Settings.Default["title"]);
                    return;
                }
                if (PesoAv.Text == "" && ModuloID.Text == "0") PesoAv.Text = "0";
                if (pesocf.Text == "" && ModuloID.Text == "0") pesocf.Text = "0"; 

                Connects.localConnect.ConnInit();
                string format = "yyyy-MM-dd";
                subQuery = "INSERT INTO CalendarioAV (Codigo_Curso, RefAcao, MomentoAv, ModuloID, Ordem, PesoCF, Metodologia, PesoAv, AvaliacaoN, DataI_EN, DataF_EN, DataI_ER, DataF_ER, DataI_EE, DataF_EE, DataI_EXT, DataF_EXT, Data_ER, Data_EE, Data_Ext) VALUES ('" + actualCode + "', '" + actualRef + "', '" + "AVALIAÇÃO " + avaliacao.Text + "', " + ModuloID.Text + ", " + 1000 + ", " + pesocf.Text.Replace(",", ".") + ", '" + metodologia.Text + "', ";
                bool res = int.TryParse(PesoAv.Text, out int a);

                if (res) subQuery += "'" + PesoAv.Text + "%" + "', "; 
                else subQuery += "'" + PesoAv.Text + "', ";
                subQuery += ConvertRomanToInt(avaliacao.Text) + ", ";

                if (op_datai_en.Checked == false) subQuery += "Null,";
                else subQuery += "'" + DataI_EN.Value.ToString(format) + "',";
                if (op_dataf_en.Checked == false) subQuery += "Null,";
                else subQuery += "'" + DataF_EN.Value.ToString(format) + "',"; 
                if (op_datai_er.Checked == false) subQuery += "Null,";
                else subQuery += "'" + DataI_ER.Value.ToString(format) + "',";
                if (op_dataf_er.Checked == false) subQuery += "Null,";
                else subQuery += "'" + DataF_ER.Value.ToString(format) + "',";
                if (op_datai_ee.Checked == false) subQuery += "Null,";
                else subQuery += "'" + DataI_EE.Value.ToString(format) + "',";
                if (op_dataf_ee.Checked == false) subQuery += "Null,";
                else subQuery += "'" + DataF_EE.Value.ToString(format) + "',";
                if (op_DataI_EXT.Checked == false) subQuery += "Null,";
                else subQuery += "'" + DataI_EXT.Value.ToString(format) + "',";
                if (op_DataF_EXT.Checked == false) subQuery += "Null,";
                else subQuery += "'" + DataF_EXT.Value.ToString(format) + "',";
                if (op_data_er.Checked == false) subQuery += "Null,";
                else subQuery += "'" + data_er.Value.ToString(format) + "',";
                if (op_data_ee.Checked == false) subQuery += "Null,";
                else subQuery += "'" + data_ee.Value.ToString(format) + "',";
                if (op_data_ext.Checked == false) subQuery += "Null)";
                else subQuery += "'" + data_ext.Value.ToString(format) + "')";

                SqlCommand insertModulo = new SqlCommand(subQuery, Connects.localConnect.Conn);
                insertModulo.CommandText = subQuery;
                insertModulo.ExecuteNonQuery();
                Connects.localConnect.ConnEnd();
                faz = true;
            }

            if (faz)
            {
                bool ok = false;
                if (op_data_ext.Checked == true || op_datai_en.Checked == true || op_dataf_en.Checked == true || op_data_er.Checked == true || op_datai_er.Checked == true || op_dataf_er.Checked == true || op_data_ee.Checked == true || op_datai_ee.Checked == true || op_dataf_ee.Checked == true || op_data_ext.Checked == true || op_DataI_EXT.Checked == true || op_DataF_EXT.Checked == true)  
                    ok = true;
                if (ok)
                {
                    string avaliacaotxt = "AVALIAÇÃO " + avaliacao.Text;
                    string subQuery2 = "SELECT ID FROM CalendarioAV WHERE metodologia <> 'Avaliação Contínua' and momentoav = '" + avaliacaotxt + "' AND Codigo_Curso = '" + actualCode + "' AND REFACAO = '" + actualRef + "'";
                    SqlDataAdapter adapter2 = new SqlDataAdapter(subQuery2, Connects.localConnect.Conn);
                    DataTable subData2 = new DataTable();
                    adapter2.Fill(subData2);
                    List<DataRow> avaliacoesguardadas2 = subData2.AsEnumerable().ToList();
                    Connects.HTlocalConnect.ConnEnd();

                    if (avaliacoesguardadas2.Count > 1)
                    {
                        if (MessageBox.Show("Atenção, pretende atualizar as datas em todas as metodologias da presente avaliação?", (string)Properties.Settings.Default["title"], MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                        {
                            for (int j = 0; j < avaliacoesguardadas2.Count; j++)
                            {
                                Connects.localConnect.ConnInit();
                                string format = "yyyy-MM-dd";
                                subQuery = "UPDATE CalendarioAV SET DataI_EN = @DataI_EN, DataF_EN = @DataF_EN, DataI_ER = @DataI_ER, DataF_ER = @DataF_ER, DataI_EE = @DataI_EE, DataF_EE = @DataF_EE, DataI_EXT = @DataI_EXT, DataF_EXT = @DataF_EXT, Data_ER = @Data_ER, Data_EE = @Data_EE, Data_Ext = @Data_Ext WHERE ID = @ID";

                                using (var sqlUpdate = new SqlCommand(subQuery, Connects.localConnect.Conn))
                                {
                                    sqlUpdate.Parameters.Add("@ID", SqlDbType.Int).Value = avaliacoesguardadas2[j][0];
                                    if (!op_datai_en.Checked)
                                        sqlUpdate.Parameters.Add("@DataI_EN", SqlDbType.DateTime).Value = DBNull.Value;
                                    else sqlUpdate.Parameters.Add("@DataI_EN", SqlDbType.DateTime).Value = DataI_EN.Value.ToString(format);

                                    if (op_dataf_en.Checked == false)
                                        sqlUpdate.Parameters.Add("@DataF_EN", SqlDbType.DateTime).Value = DBNull.Value;
                                    else sqlUpdate.Parameters.Add("@DataF_EN", SqlDbType.DateTime).Value = DataF_EN.Value.ToString(format);

                                    if (op_datai_er.Checked == false)
                                        sqlUpdate.Parameters.Add("@DataI_ER", SqlDbType.DateTime).Value = DBNull.Value;
                                    else sqlUpdate.Parameters.Add("@DataI_ER", SqlDbType.DateTime).Value = DataI_ER.Value.ToString(format);

                                    if (op_dataf_er.Checked == false)
                                        sqlUpdate.Parameters.Add("@DataF_ER", SqlDbType.DateTime).Value = DBNull.Value;
                                    else
                                        sqlUpdate.Parameters.Add("@DataF_ER", SqlDbType.DateTime).Value = DataF_ER.Value.ToString(format);

                                    if (op_datai_ee.Checked == false)
                                        sqlUpdate.Parameters.Add("@DataI_EE", SqlDbType.DateTime).Value = DBNull.Value;
                                    else
                                        sqlUpdate.Parameters.Add("@DataI_EE", SqlDbType.DateTime).Value = DataI_EE.Value.ToString(format);

                                    if (op_dataf_ee.Checked == false)
                                        sqlUpdate.Parameters.Add("@DataF_EE", SqlDbType.DateTime).Value = DBNull.Value;
                                    else
                                        sqlUpdate.Parameters.Add("@DataF_EE", SqlDbType.DateTime).Value = DataF_EE.Value.ToString(format);

                                    if (op_DataI_EXT.Checked == false)
                                        sqlUpdate.Parameters.Add("@DataI_EXT", SqlDbType.DateTime).Value = DBNull.Value;
                                    else
                                        sqlUpdate.Parameters.Add("@DataI_EXT", SqlDbType.DateTime).Value = DataI_EXT.Value.ToString(format);

                                    if (op_DataF_EXT.Checked == false)
                                        sqlUpdate.Parameters.Add("@DataF_EXT", SqlDbType.DateTime).Value = DBNull.Value;
                                    else
                                        sqlUpdate.Parameters.Add("@DataF_EXT", SqlDbType.DateTime).Value = DataF_EXT.Value.ToString(format);

                                    if (op_data_er.Checked == false)
                                        sqlUpdate.Parameters.Add("@Data_ER", SqlDbType.DateTime).Value = DBNull.Value;
                                    else
                                        sqlUpdate.Parameters.Add("@Data_ER", SqlDbType.DateTime).Value = data_er.Value.ToString(format);

                                    if (op_data_ee.Checked == false)
                                        sqlUpdate.Parameters.Add("@Data_EE", SqlDbType.DateTime).Value = DBNull.Value;
                                    else
                                        sqlUpdate.Parameters.Add("@Data_EE", SqlDbType.DateTime).Value = data_ee.Value.ToString(format);

                                    if (op_data_ext.Checked == false)
                                        sqlUpdate.Parameters.Add("@Data_EXT", SqlDbType.DateTime).Value = DBNull.Value;
                                    else
                                        sqlUpdate.Parameters.Add("@Data_EXT", SqlDbType.DateTime).Value = data_ext.Value.ToString(format);

                                    sqlUpdate.ExecuteNonQuery();
                                }
                                Connects.localConnect.ConnEnd();
                            }
                        }
                    }
                }
                Carrega_calendario();
            }
            Limpa_linha();
            Cursor.Current = Cursors.Default;
        }
        private void GridControl2_DoubleClick(object sender, EventArgs e)
        {
            GridView view = ((GridView)(gridControl2.MainView));
            var noticeID = view.GetRowCellValue(view.FocusedRowHandle, "ID");
            ID.Text = noticeID.ToString();

            if (ID.Text == "0")
            {
                MessageBox.Show("Deve guardar os registos para continuar com a edição.", (string)Properties.Settings.Default["title"]); 
                return;            
            }
            
            if (ID.Text != "0")
            {
                if (MessageBox.Show("Pretende editar ou eliminar o registo selecionado?", (string)Properties.Settings.Default["title"], MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    string subQuery = "SELECT * FROM CalendarioAV WHERE ID = '" + ID.Text + "'";
                    SqlDataAdapter adapter = new SqlDataAdapter(subQuery, Connects.localConnect.Conn);
                    DataTable subData = new DataTable();
                    adapter.Fill(subData);
                    List<DataRow> avaliacoesguardadas = subData.AsEnumerable().ToList();
                    Connects.HTlocalConnect.ConnEnd();

                    if (avaliacoesguardadas.Count > 0)
                    {
                        ID.Text = avaliacoesguardadas[0][0].ToString();
                        avaliacao.Text = ToRoman(int.Parse(avaliacoesguardadas[0][16].ToString()));
                        pesocf.Text = avaliacoesguardadas[0][5].ToString();
                        if (int.Parse(avaliacoesguardadas[0][3].ToString()) == 0)
                        {
                            ModuloID.Text = "0";
                            modulos.Text = "";
                            modulos.Enabled = false;
                            PesoAv.Enabled = false;
                        }
                        else
                        {
                            ModuloID.Text = modulosHT.Where(x => x.Modulo_ID == int.Parse(avaliacoesguardadas[0][3].ToString())).First().Modulo_ID.ToString();
                            modulos.Text = modulosHT.Where(x => x.Modulo_ID == int.Parse(avaliacoesguardadas[0][3].ToString())).First().Descricao;
                            modulos.Enabled = true;
                            PesoAv.Enabled = true;
                        }
                        metodologia.Text = avaliacoesguardadas[0][6].ToString();
                        PesoAv.Text = avaliacoesguardadas[0][7].ToString();
                        if (avaliacoesguardadas[0][8].ToString() != "")
                        {
                            op_datai_en.Checked = true;
                            op_datai_en.Enabled = true;
                            DataI_EN.Value = (DateTime)avaliacoesguardadas[0][8];
                            DataI_EN.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            DataI_EN.Enabled = false;
                            op_datai_en.Checked = false;
                        }

                        if (avaliacoesguardadas[0][9].ToString() != "")
                        {
                            op_dataf_en.Checked = true;
                            op_dataf_en.Enabled = true;
                            DataF_EN.Value = (DateTime)avaliacoesguardadas[0][9];
                            DataF_EN.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            DataF_EN.Enabled = false;
                            op_dataf_en.Checked = false;
                        }

                        if (avaliacoesguardadas[0][10].ToString() != "")
                        {
                            op_datai_er.Checked = true;
                            op_datai_er.Enabled = true;
                            DataI_ER.Value = (DateTime)avaliacoesguardadas[0][10];
                            DataI_ER.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            DataI_ER.Enabled = false;
                            op_datai_er.Checked = false;
                        }
                        if (avaliacoesguardadas[0][11].ToString() != "")
                        {
                            op_dataf_er.Checked = true;
                            op_dataf_er.Enabled = true;
                            DataF_ER.Value = (DateTime)avaliacoesguardadas[0][11];
                            DataF_ER.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            DataF_ER.Enabled = false;
                            op_dataf_er.Checked = false;
                        }
                        if (avaliacoesguardadas[0][12].ToString() != "")
                        {
                            op_datai_ee.Checked = true;
                            op_datai_ee.Enabled = true;
                            DataI_EE.Value = (DateTime)avaliacoesguardadas[0][12];
                            DataI_EE.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            DataI_EE.Enabled = false;
                            op_datai_ee.Checked = false;
                        }
                        if (avaliacoesguardadas[0][13].ToString() != "")
                        {
                            op_dataf_ee.Checked = true;
                            op_dataf_ee.Enabled = true;
                            DataF_EE.Value = (DateTime)avaliacoesguardadas[0][13];
                            DataF_EE.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            DataF_EE.Enabled = false;
                            op_dataf_ee.Checked = false;
                        }
                        if (avaliacoesguardadas[0][14].ToString() != "")
                        {
                            op_DataI_EXT.Checked = true;
                            op_DataI_EXT.Enabled = true;
                            DataI_EXT.Value = (DateTime)avaliacoesguardadas[0][14];
                            DataI_EXT.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            DataI_EXT.Enabled = false;
                            op_DataI_EXT.Checked = false;
                        }
                        if (avaliacoesguardadas[0][15].ToString() != "")
                        {
                            op_DataF_EXT.Checked = true;
                            op_DataF_EXT.Enabled = true;
                            DataF_EXT.Value = (DateTime)avaliacoesguardadas[0][15];
                            DataF_EXT.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            DataF_EXT.Enabled = false;
                            op_DataF_EXT.Checked = false;
                        }
                        if (avaliacoesguardadas[0][18].ToString() != "")
                        {
                            op_data_er.Checked = true;
                            op_data_er.Enabled = true;
                            data_er.Value = (DateTime)avaliacoesguardadas[0][18];
                            data_er.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            data_er.Enabled = false;
                            op_data_er.Checked = false;
                        }
                        if (avaliacoesguardadas[0][19].ToString() != "")
                        {
                            op_data_ee.Checked = true;
                            op_data_ee.Enabled = true;
                            data_ee.Value = (DateTime)avaliacoesguardadas[0][19];
                            data_ee.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            data_ee.Enabled = false;
                            op_data_ee.Checked = false;
                        }
                        if (avaliacoesguardadas[0][20].ToString() != "")
                        {
                            op_data_ext.Checked = true;
                            op_data_ext.Enabled = true;
                            data_ext.Value = (DateTime)avaliacoesguardadas[0][20];
                            data_ext.Format = DateTimePickerFormat.Short;
                        }
                        else
                        {
                            data_ext.Enabled = false;
                            op_data_ext.Checked = false;
                        }
                    }
                }
            }
        }
        private void SimpleButton2_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (ID.Text != "0")
            {
                if (MessageBox.Show("Atenção, pretende eliminar o registo selecionado?", (string)Properties.Settings.Default["title"], MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    Connects.localConnect.ConnInit();
                    string subQuery = "DELETE FROM CalendarioAV WHERE ID = '" + ID.Text + "'";
                    SqlCommand insertModulo = new SqlCommand(subQuery, Connects.localConnect.Conn);
                    insertModulo.CommandText = subQuery;
                    insertModulo.ExecuteNonQuery();
                    Connects.localConnect.ConnEnd();
                    Limpa_linha();
                    Carrega_calendario();
                }
            }
            Cursor.Current = Cursors.Default;
        }
        private void Limpa_linha()
        {
            pesocf.Text = "";
            avaliacao.Text = "";
            ID.Text = "";
            PesoAv.Text = "";
            metodologia.Text = "";
            ModuloID.Text = "";
            modulos.Text = "";
            modulos.Enabled = true;

            op_datai_en.Checked = false;
            op_dataf_en.Checked = false;
            DataI_EN.Enabled = false;
            DataF_EN.Enabled = false;

            op_data_er.Checked = false;
            op_datai_er.Checked = false;
            op_dataf_er.Checked = false;
            DataI_ER.Enabled = false;
            DataF_ER.Enabled = false;
            data_er.Enabled = false;

            op_data_ee.Checked = false;
            op_datai_ee.Checked = false;
            op_dataf_ee.Checked = false;
            DataI_EE.Enabled = false;
            DataF_EE.Enabled = false;
            data_ee.Enabled = false;

            op_data_ext.Checked = false;
            op_DataI_EXT.Checked = false;
            op_DataF_EXT.Checked = false;
            data_ext.Enabled = false;
            DataI_EXT.Enabled = false;
            DataF_EXT.Enabled = false;
        }
        private void Avaliacao_Leave(object sender, EventArgs e)
        {
            if (avaliacao.Text == "M")
            {
                modulos.Enabled = false; modulos.Text = "";
                ModuloID.Text = "0"; PesoAv.Enabled = false;
                op_datai_en.Checked = false; op_dataf_en.Checked = false; 
                DataI_EN.Enabled = false; DataF_EN.Enabled = false;
                op_data_er.Checked = false; op_datai_er.Checked = false;
                op_dataf_er.Checked = false; DataI_ER.Enabled = false;
                DataF_ER.Enabled = false; data_er.Enabled = false;
                op_data_ee.Checked = false; op_datai_ee.Checked = false;
                op_dataf_ee.Checked = false; DataI_EE.Enabled = false;
                DataF_EE.Enabled = false; data_ee.Enabled = false;
                op_datai_en.Enabled = false; op_dataf_en.Enabled = false;
                op_data_er.Enabled = false; op_datai_er.Enabled = false;
                op_dataf_er.Enabled = false; op_data_ee.Enabled = false; 
                op_datai_ee.Enabled = false; op_dataf_ee.Enabled = false;
                op_data_ext.Enabled = true; op_DataI_EXT.Enabled = true;
                op_DataF_EXT.Enabled = true; op_data_ext.Checked = true;
                op_DataI_EXT.Checked = true; op_DataF_EXT.Checked = true;
            }
            else
            {
                op_datai_en.Enabled = true; op_dataf_en.Enabled = true;
                op_data_er.Enabled = true; op_datai_er.Enabled = true;
                op_dataf_er.Enabled = true; op_data_ee.Enabled = true;
                op_datai_ee.Enabled = true; op_dataf_ee.Enabled = true;
                op_datai_en.Checked = true; op_dataf_en.Checked = true;
                op_data_er.Checked = true; op_datai_er.Checked = true;
                op_dataf_er.Checked = true; op_data_ee.Checked = true;
                op_datai_ee.Checked = true; op_dataf_ee.Checked = true;
                op_data_ext.Checked = false; op_DataI_EXT.Checked = false;
                op_DataF_EXT.Checked = false; op_data_ext.Enabled = false;
                op_DataI_EXT.Enabled = false; op_DataF_EXT.Enabled = false;
            }
        }
        private void Op_datai_en_CheckedChanged(object sender, EventArgs e)
        {
            DataI_EN.Enabled = op_datai_en.Checked;
        }
        private void Op_dataf_en_CheckedChanged(object sender, EventArgs e)
        {
            DataF_EN.Enabled = op_dataf_en.Checked;
        }
        private void Op_data_er_CheckedChanged(object sender, EventArgs e)
        {
            data_er.Enabled = op_data_er.Checked;
        }
        private void Op_datai_er_CheckedChanged(object sender, EventArgs e)
        {
            DataI_ER.Enabled = op_datai_er.Checked;
        }
        private void Op_dataf_er_CheckedChanged(object sender, EventArgs e)
        {
            DataF_ER.Enabled = op_dataf_er.Checked;
        }
        private void Op_data_ee_CheckedChanged(object sender, EventArgs e)
        {
            data_ee.Enabled = op_data_ee.Checked;
        }
        private void Op_datai_ee_CheckedChanged(object sender, EventArgs e)
        {
            DataI_EE.Enabled = op_datai_ee.Checked;
        }
        private void Op_dataf_ee_CheckedChanged(object sender, EventArgs e)
        {
            DataF_EE.Enabled = op_dataf_ee.Checked;
        }
        private void Op_data_ext_CheckedChanged(object sender, EventArgs e)
        {
            data_ext.Enabled = op_data_ext.Checked;
        }
        private void Op_DataI_EXT_CheckedChanged(object sender, EventArgs e)
        {
            DataI_EXT.Enabled = op_DataI_EXT.Checked;
        }
        private void Op_DataF_EXT_CheckedChanged(object sender, EventArgs e)
        {
            DataF_EXT.Enabled = op_DataF_EXT.Checked;
        }
    }
}