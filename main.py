import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick

def main():
    # ---------------------------------------------------------
    # Configuración
    # ---------------------------------------------------------
    input_file = "train.csv"
    output_csv = "resultado.csv"
    output_png = "grafico.png"
    output_excel = "reporte_ventas.xlsx"

    required_columns = ['Order Date', 'Category', 'Sales']

    # ---------------------------------------------------------
    # 1. Validación de archivo
    # ---------------------------------------------------------
    if not os.path.exists(input_file):
        print(f"❌ No se encontró el archivo '{input_file}'")
        return

    try:
        df = pd.read_csv(input_file)
    except Exception as e:
        print(f"❌ Error al leer el archivo: {e}")
        return

    # Validar columnas
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        print(f"❌ Faltan columnas: {missing_cols}")
        return

    # ---------------------------------------------------------
    # 2. Limpieza de datos
    # ---------------------------------------------------------
    df = df.dropna(subset=required_columns)

    df['Order Date'] = pd.to_datetime(df['Order Date'], dayfirst=True, errors='coerce')

    if df['Sales'].dtype == 'object':
        df['Sales'] = df['Sales'].astype(str).str.replace(r'[^\d.-]', '', regex=True)

    df['Sales'] = pd.to_numeric(df['Sales'], errors='coerce')

    df = df.dropna(subset=['Order Date', 'Sales'])

    if df.empty:
        print("⚠️ No hay datos válidos")
        return

    # ---------------------------------------------------------
    # 3. Análisis
    # ---------------------------------------------------------
    sales_category = df.groupby('Category')['Sales'].sum().sort_values(ascending=False)
    sales_category_df = sales_category.reset_index()

    total_sales = df['Sales'].sum()
    avg_sales = df['Sales'].mean()

    top_category = sales_category.idxmax()
    top_sales = sales_category.max()

    df['Month'] = df['Order Date'].dt.to_period('M')
    sales_month = df.groupby('Month')['Sales'].sum().reset_index()

    # ---------------------------------------------------------
    # 4. Guardar CSV
    # ---------------------------------------------------------
    sales_category.to_csv(output_csv)

    # ---------------------------------------------------------
    # 5. Gráfico PRO
    # ---------------------------------------------------------
    try:
        top_n = sales_category.head(10)

        plt.figure()

        ax = top_n.plot(kind='bar')

        ax.yaxis.set_major_formatter(mtick.StrMethodFormatter('${x:,.0f}'))

        for i, v in enumerate(top_n):
            plt.text(i, v, f"${v:,.0f}", ha='center', va='bottom')

        plt.title(f'Top {len(top_n)} Categorías por Ventas')
        plt.ylabel('Ventas ($)')
        plt.xticks(rotation=0)

        plt.tight_layout()
        plt.savefig(output_png)
        plt.close()

        print(f"📊 Gráfico generado: {output_png}")

    except Exception as e:
        print(f"❌ Error en gráfico: {e}")

    # ---------------------------------------------------------
    # 6. Excel PROFESIONAL
    # ---------------------------------------------------------
    try:
        with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:

            # ---------------- KPIs ----------------
            kpi_df = pd.DataFrame({
                'Metric': ['Total Sales', 'Average Sales', 'Top Category'],
                'Value': [total_sales, avg_sales, top_category]
            })

            kpi_df.to_excel(writer, sheet_name="KPIs", index=False)

            # ---------------- Resumen ----------------
            sales_category_df.to_excel(writer, sheet_name="Resumen", index=False)

            # ---------------- Ventas por Mes ----------------
            sales_month.to_excel(writer, sheet_name="Ventas por Mes", index=False)

            # ---------------- Datos Limpios ----------------
            df.to_excel(writer, sheet_name="Datos Limpios", index=False)

            # ---------------- Insights ----------------
            insights = pd.DataFrame({
                "Insight": [
                    f"La categoría con más ventas es {top_category}",
                    f"Ventas totales alcanzan ${total_sales:,.2f}",
                    f"Promedio por orden es ${avg_sales:,.2f}"
                ]
            })

            insights.to_excel(writer, sheet_name="Insights", index=False)

            # -------------------------------------------------
            # FORMATO EXCEL
            # -------------------------------------------------
            workbook = writer.book

            money_format = workbook.add_format({'num_format': '$#,##0.00'})
            bold_format = workbook.add_format({'bold': True})

            # KPIs formato
            worksheet_kpi = writer.sheets["KPIs"]
            worksheet_kpi.set_column(1, 1, 20, money_format)
            worksheet_kpi.write('A1', '📊 Resumen Ejecutivo', bold_format)

            # Tabla Resumen
            worksheet_resumen = writer.sheets["Resumen"]
            (rows_r, cols_r) = sales_category_df.shape

            worksheet_resumen.add_table(0, 0, rows_r, cols_r - 1, {
                'columns': [{'header': col} for col in sales_category_df.columns],
                'style': 'Table Style Medium 9'
            })

            # Tabla Datos Limpios
            worksheet_data = writer.sheets["Datos Limpios"]
            (rows_d, cols_d) = df.shape

            worksheet_data.add_table(0, 0, rows_d, cols_d - 1, {
                'columns': [{'header': col} for col in df.columns],
                'style': 'Table Style Medium 9'
            })

            # Ajustar columnas automáticamente
            for sheet_name, data in {
                "Resumen": sales_category_df,
                "Datos Limpios": df
            }.items():
                worksheet = writer.sheets[sheet_name]
                for i, col in enumerate(data.columns):
                    column_len = max(data[col].astype(str).map(len).max(), len(str(col)))
                    worksheet.set_column(i, i, column_len + 2)

        print(f"📁 Excel generado: {output_excel}")

    except ImportError:
        print("❌ Instala xlsxwriter: pip install xlsxwriter")
    except Exception as e:
        print(f"❌ Error en Excel: {e}")

    # ---------------------------------------------------------
    # 7. Output final
    # ---------------------------------------------------------
    print("\n✅ PROCESO COMPLETADO")
    print(f"💰 Ventas totales: ${total_sales:,.2f}")
    print(f"🏆 Mejor categoría: {top_category}")

if __name__ == "__main__":
    main()