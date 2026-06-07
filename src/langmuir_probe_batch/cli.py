"""Command-line interface for Langmuir probe batch processing."""

from __future__ import annotations

import argparse
import shutil
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd

from . import core
from .constants import (
    ARGON_ION_MASS_KG,
    ELECTRON_MASS_KG,
    ELEMENTARY_CHARGE_C,
    DEFAULT_PROBE_DIAMETER_M,
    cylindrical_probe_area,
)


@dataclass
class AnalysisConfig:
    input_dir: Path
    output_dir: Path
    shape: str = "cylindrical"
    file_pattern: str = "*.xlsx"
    probe_diameter_m: float = DEFAULT_PROBE_DIAMETER_M
    probe_area_m2: float | None = None
    ion_mass_kg: float = ARGON_ION_MASS_KG
    plot: bool = False
    remove_bad_data: bool = False

    @property
    def probe_radius_m(self) -> float:
        return self.probe_diameter_m / 2

    @property
    def area_m2(self) -> float:
        return self.probe_area_m2 or cylindrical_probe_area(diameter_m=self.probe_diameter_m)


def _find_files(input_dir: Path, pattern: str) -> list[str]:
    files = sorted(str(path) for path in input_dir.glob(pattern) if path.is_file())
    if not files:
        raise FileNotFoundError(f"No files matching {pattern!r} found in {input_dir}")
    return files


def _remove_file_indices(files: list[str], indices: list[int], bad_data_dir: Path, *, remove: bool) -> list[str]:
    if not indices:
        return files
    bad_data_dir.mkdir(parents=True, exist_ok=True)
    keep: list[str] = []
    remove_set = set(indices)
    for idx, file in enumerate(files):
        path = Path(file)
        if idx in remove_set:
            target = bad_data_dir / path.name
            if path.exists() and path.resolve() != target.resolve():
                shutil.copy2(path, target)
                if remove:
                    path.unlink()
        else:
            keep.append(str(path))
    return keep


def _filter_by_indices(values: list, indices_to_remove: list[int]) -> list:
    remove_set = set(indices_to_remove)
    return [value for idx, value in enumerate(values) if idx not in remove_set]


def _write_summary_xlsx(output_path: Path, files: list[str], vf, vp, isat, et, density) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(
        {
            "file": [Path(f).name for f in files],
            "floating_potential_v": vf,
            "plasma_potential_v": vp,
            "ion_saturation_current_mA": isat,
            "electron_temperature_eV": et,
            "plasma_density_m^-3": density,
        }
    )
    df.to_excel(output_path, index=False)


def run_analysis(config: AnalysisConfig) -> Path:
    """Run the non-iterated Langmuir probe analysis workflow."""
    config.output_dir.mkdir(parents=True, exist_ok=True)
    bad_data_dir = config.output_dir / "bad_data"

    files = _find_files(config.input_dir, config.file_pattern)

    dataframes = core.BatchRd(files)
    # Keep the original bad-data function available in core.py, but avoid calling it here because
    # it deletes input files and assumes a specific folder layout. This CLI quarantines bad files.
    bad_indices = []
    for idx, df in enumerate(dataframes):
        if "Current" not in df.columns:
            bad_indices.append(idx)
            continue
        current = df.loc[:, "Current"].to_numpy()
        if len(np.where(current > 0)[0]) == 0 or all(current < 1e-5):
            bad_indices.append(idx)
    files = _remove_file_indices(files, bad_indices, bad_data_dir, remove=config.remove_bad_data)

    current, voltage = core.VC(files)
    zero_cross = core.ZeroCross(files, current)
    current_ranges, voltage_ranges = core.ZCAnalysisRng(zero_cross, current, voltage)
    vfloat = core.FloatingPot(np.array(current_ranges), np.array(voltage_ranges))

    complex_indices = [idx for idx, value in enumerate(vfloat) if not isinstance(value, np.floating)]
    files = _remove_file_indices(files, complex_indices, bad_data_dir, remove=config.remove_bad_data)
    if complex_indices:
        vfloat = np.array(_filter_by_indices(vfloat, complex_indices), dtype=np.float64)
        current = _filter_by_indices(current, complex_indices)
        voltage = _filter_by_indices(voltage, complex_indices)
        zero_cross = _filter_by_indices(zero_cross, complex_indices)

    electron_current, _, _ = core.ECurr(vfloat, files, current, voltage)
    vplasma = core.PlasmaPot(files, zero_cross, electron_current, voltage)
    electron_temp, r_squared = core.ElectronTemp(
        files, electron_current, vfloat, vplasma, voltage, config.ion_mass_kg, ELECTRON_MASS_KG
    )

    negative_temp_indices = [idx for idx, value in enumerate(electron_temp) if value < 0]
    files = _remove_file_indices(files, negative_temp_indices, bad_data_dir, remove=config.remove_bad_data)
    if negative_temp_indices:
        electron_temp = _filter_by_indices(electron_temp, negative_temp_indices)
        current = _filter_by_indices(current, negative_temp_indices)
        voltage = _filter_by_indices(voltage, negative_temp_indices)
        zero_cross = _filter_by_indices(zero_cross, negative_temp_indices)
        vfloat = np.array(_filter_by_indices(vfloat, negative_temp_indices), dtype=np.float64)
        vplasma = _filter_by_indices(vplasma, negative_temp_indices)
        electron_current = _filter_by_indices(electron_current, negative_temp_indices)

    plasma_props = core.PlasmaProps(
        config.shape,
        ELEMENTARY_CHARGE_C,
        config.area_m2,
        config.probe_radius_m,
        vfloat,
        vplasma,
        electron_current,
        config.ion_mass_kg,
        ELECTRON_MASS_KG,
        electron_temp,
        zero_cross,
        files,
        current,
        voltage,
    )

    density = plasma_props[0]
    ion_saturation_current = plasma_props[2]
    output_path = config.output_dir / "lp_data_post_processing_summary.xlsx"
    _write_summary_xlsx(output_path, files, vfloat, vplasma, ion_saturation_current, electron_temp, density)

    if config.plot:
        core.PlotData(files, zero_cross, voltage_ranges, current_ranges, vplasma, vfloat)

    print(f"Processed {len(files)} files")
    print(f"Electron-temperature fit R^2 values: {r_squared}")
    print(f"Summary written to: {output_path}")
    return output_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Batch-process Langmuir probe I-V sweep files.")
    parser.add_argument("input_dir", type=Path, help="Folder containing raw Langmuir probe Excel files.")
    parser.add_argument("-o", "--output-dir", type=Path, default=Path("results"), help="Output directory.")
    parser.add_argument("--pattern", default="*.xlsx", help="Input glob pattern. Default: *.xlsx")
    parser.add_argument("--shape", default="cylindrical", choices=["cylindrical", "spherical", "planar"], help="Probe shape.")
    parser.add_argument("--probe-diameter-m", type=float, default=DEFAULT_PROBE_DIAMETER_M, help="Probe diameter in meters.")
    parser.add_argument("--probe-area-m2", type=float, default=None, help="Override probe area in square meters.")
    parser.add_argument("--ion-mass-kg", type=float, default=ARGON_ION_MASS_KG, help="Ion mass in kg. Default: argon.")
    parser.add_argument("--plot", action="store_true", help="Generate diagnostic plots.")
    parser.add_argument("--remove-bad-data", action="store_true", help="Delete bad input files after copying them to output/bad_data.")
    return parser


def main(argv: list[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    config = AnalysisConfig(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        shape=args.shape,
        file_pattern=args.pattern,
        probe_diameter_m=args.probe_diameter_m,
        probe_area_m2=args.probe_area_m2,
        ion_mass_kg=args.ion_mass_kg,
        plot=args.plot,
        remove_bad_data=args.remove_bad_data,
    )
    run_analysis(config)


if __name__ == "__main__":
    main()
