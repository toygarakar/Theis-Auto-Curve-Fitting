# Theis-Auto-Curve-Fitting
An Automated Tool for Estimating Aquifer Parameters from Pumping-Test Data

**Theis Auto Curve Fitting** is an Excel VBA application that automatically fits the Theis (1935) analytical solution to time–drawdown data from pumping tests to estimate aquifer hydraulic parameters.

The application automates the curve-fitting process through iterative optimization, providing numerical estimates and graphical comparisons between observed drawdown data and the fitted Theis curve.

## Features

- **Automated curve fitting:** Fits the Theis analytical solution to observed time–drawdown data.
- **Parameter estimation:** Estimates transmissivity \(T\) and storativity \(S\) through iterative optimization.
- **Fit assessment:** Reports numerical error metrics, including weighted and unweighted sum of squared errors (SSE) and (R<sup>2</sup>), as implemented in the application.
- **Graphical visualization:** Displays observed drawdown measurements as blue markers and the fitted Theis curve as a red line.
- **Excel-based workflow:** Performs the analysis within Microsoft Excel using VBA, without requiring a separate numerical modelling environment.

## Theoretical Background

The application is based on the analytical solution developed by Theis (1935) for transient groundwater flow toward a pumping well in a confined aquifer.

The Theis equation is expressed as

$$
s(r,t)=\frac{Q}{4\pi T}W(u)
$$

where

$$
u=\frac{r^2S}{4Tt}
$$

and:

- \(s\) = drawdown  (L)
- \(Q\) = pumping rate (L<sup>3</sup>/T)
- \(T\) = transmissivity (L<sup>2</sup>/T)
- \(S\) = storativity (dimensionless)
- \(r\) = radial distance from the pumping well to the observation point (L)
- \(t\) = elapsed time (t)
- \(W(u)\) = Theis well function, defined by the exponential integral (dimensionless)

The application estimates the aquifer parameters by fitting the theoretical drawdown response to observed pumping-test data.

Hydraulic conductivity (\(K\)) can be derived from transmissivity when the saturated aquifer thickness (\(b\)) is known:

$$
K=\frac{T}{b}
$$

## Input Data

The application uses time–drawdown measurements entered into an Excel worksheet.

| Variable | Description | Example unit |
|---|---|---|
| Time (\(t\)) | Elapsed time since pumping began | min |
| Drawdown (\(s\)) | Observed groundwater-level decline | m |
| Pumping Time (\(t\)) | Total pumping time  | min |
| Pumping Rate (\(Q\)) | The constant volumetric discharge | m<sup>3</sup>/min |
| Aquifer Thickness (\(m\)) | If the aquifer thickness is known, \(K\) can be calculated | m |
| r (\(m\)) | Radial distance from the pumping well  | m |

The pumping rate, observation-well distance, and other required model settings must be specified as required by the application.

### Example

| Time (min) | Drawdown (m) |
|---:|---:|
| 1 | 0.25 |
| 2 | 0.40 |
| 5 | 0.62 |
| 10 | 0.81 |

*The values above are illustrative and do not represent an actual pumping test.*

## How to Use

1. Open the Excel workbook containing the VBA application.
2. Enter the observed time–drawdown data in the designated worksheet.
3. Specify the required pumping-test parameters and verify that all input units are consistent.
4. Click the **Curve Fit** button to run the optimization procedure.
5. Review the estimated parameters, numerical fit statistics, and graphical comparison.

## Outputs

The application provides the following results, subject to the implemented functionality:

- **Transmissivity (\(T\))**: Estimated aquifer transmissivity.
- **Storativity (\(S\))**: Estimated aquifer storage coefficient.
- **Weighted sum of squared errors (\(SSE_W\))**: Weighted measure of the discrepancy between observed and modelled drawdown values.
- **Sum of squared errors (\(SSE\))**: Unweighted sum of squared residuals.
- **Coefficient of determination (\(R^2\))**: A statistical measure of fit, calculated according to the implementation used in the application.
- **Fitted curve plot**: A graphical comparison of observed data and the theoretical Theis response.

The precise interpretation of each fit statistic depends on its mathematical definition and implementation.

## Scope and Limitations

The Theis analytical solution assumes an idealized aquifer system, including a homogeneous, isotropic, laterally extensive confined aquifer, fully penetrating pumping and observation wells, and constant pumping rate.

Departures from these assumptions may affect parameter estimates. Early-time observations can also be influenced by wellbore storage and other non-ideal effects.

Estimated parameters should therefore be interpreted in the context of the hydrogeological setting, data quality, and assumptions of the analytical model.

## Requirements

- Microsoft Excel
- VBA macro execution enabled
- A compatible version of the application workbook

Additional requirements, supported Excel versions, and input conventions should be documented for the released version.

## Citation

If you use this software in your research, please cite the associated software publication when available.

Citation information will be provided in the `CITATION.cff` file and updated following publication.

## License

This project is distributed under the MIT License. See the `LICENSE` file for the full license text.

## Reference

Theis, C. V. (1935). The relation between the lowering of the piezometric surface and the rate and duration of discharge of a well using ground-water storage. *Eos, Transactions American Geophysical Union, 16*(2), 519–524. https://doi.org/10.1029/TR016i002p00519
