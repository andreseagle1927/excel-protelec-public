from pathlib import Path
from typing import Dict, List

import sys

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "clean"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.cleaning.normalize import normalize_spaces
from src.parsers.cuentas import parse_cuentas
from src.parsers.movimientos import parse_movimientos
from src.parsers.saldos import parse_saldos


def _clean_text_cols(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for col in cols:
        if col in df.columns:
            df[col] = df[col].fillna("").map(normalize_spaces)
    return df


def _save(df: pd.DataFrame, name: str) -> None:
    parquet_path = OUT / f"{name}.parquet"
    csv_path = OUT / f"{name}.csv"
    df.to_parquet(parquet_path, index=False)
    df.to_csv(csv_path, index=False)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)

    quality: List[Dict] = []
    rejected_rows: List[Dict] = []

    cuentas_file = ROOT / "CUENTA MATERIALES.xlsx"
    cuentas_records, cuentas_rejected = parse_cuentas(str(cuentas_file))
    dim_cuentas = pd.DataFrame(cuentas_records)
    dim_cuentas = _clean_text_cols(dim_cuentas, ["cuenta", "nombre_cuenta"])
    dim_cuentas = dim_cuentas.drop_duplicates(subset=["cuenta"], keep="first")
    _save(dim_cuentas, "dim_cuentas")
    rejected_rows.extend(cuentas_rejected)
    quality.append(
        {
            "dataset": "dim_cuentas",
            "filas": len(dim_cuentas),
            "rechazadas": len(cuentas_rejected),
            "duplicados": int(dim_cuentas.duplicated(subset=["cuenta"]).sum()),
            "nulos_criticos": int(dim_cuentas["cuenta"].eq("").sum()),
        }
    )

    mov_files = sorted(ROOT.glob("Movimiento contable Materiales *.xlsx"))
    mov_records = []
    mov_rejected = []
    for file_path in mov_files:
        records, rejected = parse_movimientos(str(file_path))
        mov_records.extend(records)
        mov_rejected.extend(rejected)

    fact_movimientos = pd.DataFrame(mov_records)
    if not fact_movimientos.empty:
        fact_movimientos = _clean_text_cols(
            fact_movimientos,
            [
                "no_oper",
                "tipo_doc",
                "no_doc",
                "clasif",
                "detalle_operacion",
                "cuenta",
                "tercero_id",
                "centro_costos",
                "detalle",
            ],
        )
        fact_movimientos = fact_movimientos.merge(
            dim_cuentas[["cuenta", "nombre_cuenta"]],
            on="cuenta",
            how="left",
            suffixes=("", "_dim"),
        )
    _save(fact_movimientos, "fact_movimientos")
    rejected_rows.extend(mov_rejected)
    quality.append(
        {
            "dataset": "fact_movimientos",
            "filas": len(fact_movimientos),
            "rechazadas": len(mov_rejected),
            "duplicados": int(
                fact_movimientos.duplicated(
                    subset=["anio", "no_oper", "cuenta", "tercero_id", "debito", "credito", "detalle"]
                ).sum()
            )
            if not fact_movimientos.empty
            else 0,
            "nulos_criticos": int(
                fact_movimientos[["cuenta", "tercero_id"]].isna().sum().sum()
            )
            if not fact_movimientos.empty
            else 0,
        }
    )

    sal_files = sorted(ROOT.glob("Saldos de terceros cuenta Materiales *.xlsx"))
    sal_records = []
    sal_rejected = []
    for file_path in sal_files:
        records, rejected = parse_saldos(str(file_path))
        sal_records.extend(records)
        sal_rejected.extend(rejected)

    fact_saldos = pd.DataFrame(sal_records)
    if not fact_saldos.empty:
        fact_saldos = _clean_text_cols(
            fact_saldos,
            ["cuenta", "nombre_cuenta", "tercero_id", "tercero_nombre", "fecha_corte"],
        )
        fact_saldos = fact_saldos.merge(
            dim_cuentas[["cuenta", "nombre_cuenta"]],
            on="cuenta",
            how="left",
            suffixes=("", "_dim"),
        )
        fact_saldos["nombre_cuenta"] = fact_saldos["nombre_cuenta_dim"].fillna(
            fact_saldos["nombre_cuenta"]
        )
        fact_saldos = fact_saldos.drop(columns=["nombre_cuenta_dim"])

    _save(fact_saldos, "fact_saldos_tercero")

    dim_terceros = pd.DataFrame(columns=["tercero_id", "tercero_nombre", "anio_referencia", "fecha_corte_referencia"])
    if not fact_saldos.empty:
        tmp_ter = fact_saldos[["tercero_id", "tercero_nombre", "anio", "fecha_corte"]].copy()
        tmp_ter = tmp_ter[tmp_ter["tercero_id"].astype(str).str.strip() != ""]
        tmp_ter["tercero_nombre"] = tmp_ter["tercero_nombre"].fillna("").astype(str).str.strip()
        tmp_ter["tercero_id"] = tmp_ter["tercero_id"].astype(str).str.strip()

        grouped = (
            tmp_ter.groupby(["tercero_id", "tercero_nombre"], dropna=False)
            .agg(
                freq=("tercero_nombre", "size"),
                anio_max=("anio", "max"),
                fecha_corte_max=("fecha_corte", "max"),
            )
            .reset_index()
            .sort_values(["tercero_id", "freq", "anio_max", "fecha_corte_max"], ascending=[True, False, False, False])
        )
        dim_terceros = grouped.drop_duplicates(subset=["tercero_id"], keep="first").rename(
            columns={"anio_max": "anio_referencia", "fecha_corte_max": "fecha_corte_referencia"}
        )[["tercero_id", "tercero_nombre", "anio_referencia", "fecha_corte_referencia"]]
        if not fact_movimientos.empty:
            fact_movimientos = fact_movimientos.merge(
                dim_terceros[["tercero_id", "tercero_nombre"]],
                on="tercero_id",
                how="left",
                suffixes=("", "_dim"),
            )
            if "tercero_nombre" not in fact_movimientos.columns and "tercero_nombre_dim" in fact_movimientos.columns:
                fact_movimientos["tercero_nombre"] = fact_movimientos["tercero_nombre_dim"]
            elif "tercero_nombre_dim" in fact_movimientos.columns:
                fact_movimientos["tercero_nombre"] = fact_movimientos["tercero_nombre"].fillna(
                    fact_movimientos["tercero_nombre_dim"]
                )
            if "tercero_nombre_dim" in fact_movimientos.columns:
                fact_movimientos = fact_movimientos.drop(columns=["tercero_nombre_dim"])

            _save(fact_movimientos, "fact_movimientos")

    _save(dim_terceros, "dim_terceros")

    rejected_rows.extend(sal_rejected)
    quality.append(
        {
            "dataset": "fact_saldos_tercero",
            "filas": len(fact_saldos),
            "rechazadas": len(sal_rejected),
            "duplicados": int(
                fact_saldos.duplicated(subset=["anio", "cuenta", "tercero_id", "saldo_actual"]).sum()
            )
            if not fact_saldos.empty
            else 0,
            "nulos_criticos": int(fact_saldos[["cuenta", "tercero_id"]].isna().sum().sum())
            if not fact_saldos.empty
            else 0,
        }
    )

    quality.append(
        {
            "dataset": "dim_terceros",
            "filas": len(dim_terceros),
            "rechazadas": 0,
            "duplicados": int(dim_terceros.duplicated(subset=["tercero_id"]).sum()) if not dim_terceros.empty else 0,
            "nulos_criticos": int(dim_terceros["tercero_id"].isna().sum()) if not dim_terceros.empty else 0,
        }
    )

    pd.DataFrame(rejected_rows).to_csv(OUT / "rejected_rows.csv", index=False)
    pd.DataFrame(quality).to_csv(OUT / "quality_summary.csv", index=False)

    print("OK: procesamiento finalizado")
    print(f"- dim_cuentas: {len(dim_cuentas)}")
    print(f"- fact_movimientos: {len(fact_movimientos)}")
    print(f"- fact_saldos_tercero: {len(fact_saldos)}")
    print(f"- dim_terceros: {len(dim_terceros)}")
    print(f"- rechazadas: {len(rejected_rows)}")


if __name__ == "__main__":
    main()
