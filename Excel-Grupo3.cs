﻿using System;
using System.Collections.Generic;
using System.Drawing; //  para trabajar con fuentes, colores, etc
using System.Globalization; //para formatos numéricos/culturales.
using System.Linq; //para ordenar, sumar, etc.
using System.Windows.Forms;

namespace Excel
{
    public partial class Form1 : Form
    {
        private bool ignorarCambioCelda = false;
        private string portapapelesTexto = "";
        private DataGridViewCell celdaCopiada = null;
        private bool operacionEnProceso = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void aggnum()
        {
            for (int i = 1; i <= 50; i++)
            {
                int rowIndex = dgvdetalle.Rows.Add();
                dgvdetalle.Rows[rowIndex].Cells["No"].Value = i;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            aggnum();
            CargarComboFuentes();

            foreach (DataGridViewRow fila in dgvdetalle.Rows)
            {
                foreach (DataGridViewCell celda in fila.Cells)
                {
                    celda.Style.Font = new Font("Calibri", 10F, FontStyle.Regular);
                }
            }
        }

        private void BTNNegrita_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                Font fuenteActual = celda.Style.Font ?? dgvdetalle.Font;
                bool esNegrita = fuenteActual.Style.HasFlag(FontStyle.Bold);
                FontStyle nuevoEstilo = esNegrita
                    ? fuenteActual.Style & ~FontStyle.Bold
                    : fuenteActual.Style | FontStyle.Bold;

                celda.Style.Font = new Font(fuenteActual.FontFamily, fuenteActual.Size, nuevoEstilo);
            }
        }

        private void BtnSubrayar_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                Font fuenteActual = celda.Style.Font ?? dgvdetalle.Font;
                bool tieneSubrayado = fuenteActual.Style.HasFlag(FontStyle.Underline);
                FontStyle nuevoEstilo = tieneSubrayado
                    ? fuenteActual.Style & ~FontStyle.Underline
                    : fuenteActual.Style | FontStyle.Underline;

                celda.Style.Font = new Font(fuenteActual.FontFamily, fuenteActual.Size, nuevoEstilo);
            }
        }

        private void BtnIzquierda_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                celda.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }
        }

        private void BtnDerecha_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                celda.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        private void BtnCentro_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                celda.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void CargarComboFuentes()
        {
            CBXFuentes.Items.Add("Calibri");
            CBXFuentes.Items.Add("Arial");
            CBXFuentes.Items.Add("Times New Roman");
            CBXFuentes.SelectedIndex = 0;
            CBXFuentes.SelectedIndexChanged += CBXFuentes_SelectedIndexChanged;

            for (int i = 8; i <= 20; i++)
            {
                comboBoxTamaño.Items.Add(i);
            }
            comboBoxTamaño.SelectedIndex = 2;
            comboBoxTamaño.SelectedIndexChanged += comboBoxTamaño_SelectedIndexChanged;

            CBXFormatoNumero.Items.Add("General");
            CBXFormatoNumero.Items.Add("L. Lempiras");
            CBXFormatoNumero.Items.Add("$ Dólares");
            CBXFormatoNumero.SelectedIndex = 0;
            CBXFormatoNumero.SelectedIndexChanged += CBXFormatoNumero_SelectedIndexChanged;

            CBXOperaciones.Items.Add("Suma");
            CBXOperaciones.Items.Add("Resta");
            CBXOperaciones.Items.Add("Multiplicación");
            CBXOperaciones.Items.Add("División");
            CBXOperaciones.SelectedIndex = 0;
            CBXOperaciones.SelectedIndexChanged += CBXOperaciones_SelectedIndexChanged;
        }

        private void CBXFuentes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ignorarCambioCelda) return;

            string fuenteSeleccionada = CBXFuentes.SelectedItem.ToString();

            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                Font fuenteActual = celda.Style.Font ?? dgvdetalle.Font;
                celda.Style.Font = new Font(fuenteSeleccionada, fuenteActual.Size, fuenteActual.Style);
            }
        }

        private void comboBoxTamaño_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ignorarCambioCelda || comboBoxTamaño.SelectedItem == null) return;

            float tamañoSeleccionado = Convert.ToSingle(comboBoxTamaño.SelectedItem);

            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                Font fuenteActual = celda.Style.Font ?? dgvdetalle.Font;
                celda.Style.Font = new Font(fuenteActual.FontFamily, tamañoSeleccionado, fuenteActual.Style);
            }
        }

        private void dgvdetalle_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (!ignorarCambioCelda)
            {
                ignorarCambioCelda = true;
                MostrarFormatoDeCelda(dgvdetalle.Rows[e.RowIndex].Cells[e.ColumnIndex]);
                ignorarCambioCelda = false;
            }
        }

        private void MostrarFormatoDeCelda(DataGridViewCell celda)
        {
            var fuente = celda.Style.Font ?? dgvdetalle.Font;
            var alineacion = celda.Style.Alignment;

            if (CBXFuentes.Items.Contains(fuente.FontFamily.Name))
                CBXFuentes.SelectedItem = fuente.FontFamily.Name;

            if (comboBoxTamaño.Items.Contains((int)fuente.Size))
                comboBoxTamaño.SelectedItem = (int)fuente.Size;

            BTNNegrita.BackColor = fuente.Bold ? Color.LightBlue : SystemColors.Control;
            BtnSubrayar.BackColor = fuente.Underline ? Color.LightBlue : SystemColors.Control;

            BtnIzquierda.BackColor = alineacion == DataGridViewContentAlignment.MiddleLeft ? Color.LightBlue : SystemColors.Control;
            BtnCentro.BackColor = alineacion == DataGridViewContentAlignment.MiddleCenter ? Color.LightBlue : SystemColors.Control;
            BtnDerecha.BackColor = alineacion == DataGridViewContentAlignment.MiddleRight ? Color.LightBlue : SystemColors.Control;

            string texto = celda.Value?.ToString() ?? "";
            if (texto.StartsWith("L."))
                CBXFormatoNumero.SelectedItem = "L. Lempiras";
            else if (texto.StartsWith("$"))
                CBXFormatoNumero.SelectedItem = "$ Dólares";
            else
                CBXFormatoNumero.SelectedItem = "General";
        }

        private void CBXFormatoNumero_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ignorarCambioCelda || CBXFormatoNumero.SelectedItem == null) return;

            string formatoSeleccionado = CBXFormatoNumero.SelectedItem.ToString();

            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                if (celda.Value == null) continue;

                string texto = celda.Value.ToString().Replace("L.", "").Replace("$", "").Trim();

                if (double.TryParse(texto, out double valor))
                {
                    string nuevoTexto = "";

                    switch (formatoSeleccionado)
                    {
                        case "L. Lempiras":
                            nuevoTexto = $"L. {valor:N2}";
                            break;
                        case "$ Dólares":
                            nuevoTexto = $"${valor:N2}";
                            break;
                        case "General":
                            nuevoTexto = valor.ToString();
                            break;
                    }

                    celda.Value = nuevoTexto;
                }
            }
        }

        private void CBXOperaciones_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (operacionEnProceso)
                return;

            if (ignorarCambioCelda)
                return;

            if (CBXOperaciones.SelectedIndex < 0)
                return;

            operacionEnProceso = true;

            try
            {
                var seleccionadas = dgvdetalle.SelectedCells
                    .Cast<DataGridViewCell>()
                    .OrderBy(c => c.RowIndex)
                    .ThenBy(c => c.ColumnIndex)
                    .ToList();

                if (seleccionadas.Count < 2)
                {
                    MessageBox.Show("Debes seleccionar al menos 2 celdas para realizar la operación.");
                    return;
                }

                var celdasConNumeros = new List<DataGridViewCell>();
                var celdasVacias = new List<DataGridViewCell>();

                foreach (var celda in seleccionadas)
                {
                    string texto = celda.Value?.ToString().Replace("L.", "").Replace("$", "").Trim() ?? "";

                    if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out double numero) ||
                        double.TryParse(texto, NumberStyles.Any, CultureInfo.CurrentCulture, out numero))
                    {
                        celdasConNumeros.Add(celda);
                    }
                    else
                    {
                        celdasVacias.Add(celda);
                    }
                }

                if (celdasConNumeros.Count == 0)
                {
                    MessageBox.Show("No hay celdas con números válidos en la selección.");
                    return;
                }

                DataGridViewCell celdaDestino = null;

                if (celdasVacias.Count > 0)
                {
                    celdaDestino = celdasVacias.First();
                }
                else
                {
                    celdaDestino = seleccionadas.Last();
                    if (celdasConNumeros.Contains(celdaDestino))
                    {
                        celdasConNumeros.Remove(celdaDestino);
                    }
                }

                if (celdasConNumeros.Count == 0)
                {
                    MessageBox.Show("No hay suficientes celdas con números para realizar la operación.");
                    return;
                }

                List<double> valores = new List<double>();
                foreach (var celda in celdasConNumeros)
                {
                    string texto = celda.Value?.ToString().Replace("L.", "").Replace("$", "").Trim() ?? "";
                    if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out double numero) ||
                        double.TryParse(texto, NumberStyles.Any, CultureInfo.CurrentCulture, out numero))
                    {
                        valores.Add(numero);
                    }
                }

                string operacion = CBXOperaciones.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(operacion))
                {
                    MessageBox.Show("Selecciona una operación válida.");
                    return;
                }

                string celdasUsadas = string.Join(", ", celdasConNumeros.Select(c =>
                    $"{GetColumnName(c.ColumnIndex)}{c.RowIndex + 1}"));

                double resultado = 0;

                switch (operacion)
                {
                    case "Suma":
                        resultado = valores.Sum();
                        break;
                    case "Resta":
                        resultado = valores[0];
                        for (int i = 1; i < valores.Count; i++)
                            resultado -= valores[i];
                        break;
                    case "Multiplicación":
                        resultado = 1;
                        foreach (var v in valores)
                            resultado *= v;
                        break;
                    case "División":
                        resultado = valores[0];


                        for (int i = 1; i < valores.Count; i++)
                        {
                            if (Math.Abs(valores[i]) < double.Epsilon) 
                            {
                                MessageBox.Show($"Error: División por cero detectada.\n\n" +
                                              $"Valores utilizados: {string.Join(", ", valores)}\n" +
                                              $"Celdas: {celdasUsadas}\n\n" +
                                              $"El valor cero está en la posición {i + 1} de la operación.\n" +
                                              $"No se puede dividir entre cero.",
                                              "Error de División",
                                              MessageBoxButtons.OK,
                                              MessageBoxIcon.Error);
                                return;
                            }
                            resultado /= valores[i];
                        }

                        if (double.IsInfinity(resultado) || double.IsNaN(resultado))
                        {
                            MessageBox.Show("Error: El resultado de la división no es válido (infinito o indeterminado).\n" +
                                          "Verifique los valores ingresados.",
                                          "Error de División",
                                          MessageBoxButtons.OK,
                                          MessageBoxIcon.Error);
                            return;
                        }
                        break;
                    default:
                        MessageBox.Show("Operación no válida.");
                        return;
                }

                string mensaje = $"Operación: {operacion}\n" +
                                $"Celdas utilizadas: {celdasUsadas}\n" +
                                $"Valores: {string.Join(", ", valores)}\n" +
                                $"Resultado: {resultado}";

                MessageBox.Show(mensaje, "Resultado de la Operación");

                celdaDestino.Value = resultado;

                dgvdetalle.ClearSelection();
                celdaDestino.Selected = true;
                dgvdetalle.CurrentCell = celdaDestino;

                ignorarCambioCelda = true;
                CBXOperaciones.SelectedIndex = 0;
                ignorarCambioCelda = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error en la operación: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                operacionEnProceso = false;
            }
        }

        private void BtnLimpiar_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell celda in dgvdetalle.SelectedCells)
            {
                celda.Value = null;
            }
        }

        private void OrdenarCeldasSeleccionadas(bool ascendente)
        {
            var celdas = dgvdetalle.SelectedCells
              .Cast<DataGridViewCell>()
              .Where(c => c.Value != null)
              .OrderBy(c => c.RowIndex)
              .ThenBy(c => c.ColumnIndex)
              .ToList();

            if (celdas.Count == 0) return;

            List<string> textos = celdas
                .Select(c => c.Value.ToString().Replace("L.", "").Replace("$", "").Trim())
                .ToList();

            // Verificar si todos son números
            bool todosNumeros = textos.All(t => double.TryParse(t, out _));

            if (todosNumeros)
            {
                // Convertir a números
                List<double> numeros = textos.Select(t => double.Parse(t)).ToList();

                if (ascendente)
                    numeros.Sort();
                else
                    numeros.Sort((a, b) => b.CompareTo(a));

                // Reasignar valores ordenados
                for (int i = 0; i < celdas.Count; i++)
                    celdas[i].Value = numeros[i];
            }
            else
            {
                // Ordenar como texto
                if (ascendente)
                    textos.Sort();
                else
                    textos.Sort((a, b) => string.Compare(b, a));

                for (int i = 0; i < celdas.Count; i++)
                    celdas[i].Value = textos[i];
            }
        }

        private void OrdenarMayor_Click(object sender, EventArgs e)
        {
            OrdenarCeldasSeleccionadas(ascendente: true);
        }

        private void OrdenarMenor_Click(object sender, EventArgs e)
        {
            OrdenarCeldasSeleccionadas(ascendente: false);
        }

        private void BtnCopiar_Click(object sender, EventArgs e)
        {
            if (dgvdetalle.CurrentCell != null && dgvdetalle.CurrentCell.Value != null)
            {
                portapapelesTexto = dgvdetalle.CurrentCell.Value.ToString();
                celdaCopiada = dgvdetalle.CurrentCell;
            }
        }

        private void BtnCortar_Click(object sender, EventArgs e)
        {
            if (dgvdetalle.CurrentCell != null && dgvdetalle.CurrentCell.Value != null)
            {
                portapapelesTexto = dgvdetalle.CurrentCell.Value.ToString();
                celdaCopiada = dgvdetalle.CurrentCell;
                dgvdetalle.CurrentCell.Value = ""; 
            }
        }

        private void BtnPegar_Click(object sender, EventArgs e)
        {
            if (dgvdetalle.CurrentCell != null && !string.IsNullOrEmpty(portapapelesTexto))
            {
                dgvdetalle.CurrentCell.Value = portapapelesTexto;
            }
        }

        private string GetColumnName(int columnIndex)
        {
            if (columnIndex == 0) return "No";

            char letra = (char)('A' + columnIndex - 1);
            return letra.ToString();
        }
    }
}