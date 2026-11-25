
# 33/11 kV Substation Design

## About the Project

This project focuses on the design and engineering analysis of a 33/11 kV distribution substation, covering the major electrical design considerations involved in power distribution, from load assessment and fault calculations to equipment selection, earthing analysis and protection coordination.

The objective of this project was to develop a structured understanding of how the key components of a distribution substation are selected, how electrical parameters influence design decisions, and how these calculations can be organised into a practical engineering workflow.

The design is based on an assumed electrical network and is developed as an independent learning and engineering practice project.

## Substation Configuration

The proposed substation consists of two incoming 33 kV lines, two 10 MVA, 33/11 kV power transformers and six outgoing 11 kV feeders.

The 11 kV bus is divided into two sections, with each transformer supplying its respective bus section. The bus couplers are normally kept open to maintain the intended operating configuration.

This arrangement forms the basis for the load assessment, fault analysis, equipment selection and protection studies carried out in this project.

## Design Approach and Methodology

I approached the project by breaking the substation design into individual engineering stages, starting with the network configuration and assumed operating parameters, followed by calculations and equipment selection.

### 1. Load Assessment and Transformer Capacity

I began by defining the feeder loads and estimating the overall demand on the substation. I then evaluated the transformer loading condition and considered future load growth at an assumed annual growth rate of 5%.

This analysis helped assess the available transformer capacity under the initial operating conditions and identify when the existing capacity would require further planning.

### 2. Short-Circuit Analysis

I calculated the prospective short-circuit current at the 33 kV and 11 kV levels using the assumed network fault level and transformer impedance.

These calculations were used to understand the fault conditions at different voltage levels and establish the basis for selecting suitable switchgear fault ratings.

### 3. Electrical Equipment Selection

Based on the calculated electrical parameters, I evaluated the ratings of major substation equipment, including circuit breakers, isolators, current transformers and associated equipment.

The objective was to relate the calculated fault conditions and operating requirements to appropriate equipment ratings.

### 4. 33 kV Busbar Selection

I considered the selection of the 33 kV busbar conductor based on the design assumptions and electrical requirements of the substation.

The selected conductor for this project is ACSR Wolf.

### 5. Earthing System Analysis

I developed an earthing grid assessment based on assumed soil resistivity, ground fault current and grid dimensions.

The analysis included:

- Earthing grid configuration
- Earth electrode arrangement
- Touch voltage assessment
- Step voltage assessment
- Ground potential rise evaluation

The results were used to assess the earthing system against the assumed safety criteria and identify areas requiring further engineering review.

### 6. Protection Coordination

I worked on the protection coordination aspect of the design by considering relay operating settings and the coordination between feeder, incomer and upstream protection.

The objective was to establish a basic protection hierarchy so that fault isolation could be considered in a coordinated manner.

### 7. Single-Line Diagram Development

After establishing the substation configuration and major equipment ratings, I developed the single-line diagram to represent the electrical arrangement of the proposed substation.

The drawing represents the incoming supply arrangement, transformers, bus sections, bus couplers and outgoing feeders.

The drawing is available in DXF, PDF and PNG formats.

## Tools and Implementation

I used Excel to organise the design calculations into a structured workbook, with separate worksheets for different calculation areas and an input sheet for the assumed parameters.

Python was used to automate the drawing generation, create the calculation workbook and independently verify selected calculation results.

The project includes:

- Excel-based engineering calculations
- Python-based automation
- Single-line diagram generation
- Calculation verification
- Structured documentation of design assumptions and methodology

## Key Design Outcomes

The design assessment produced the following results under the assumed input conditions:

- Peak load assessment: 8.24 MVA
- 33 kV short-circuit current: 13.1 kA
- 11 kV short-circuit current: 5.6 kA with one transformer and 9.8 kA with both transformers
- Selected switchgear fault ratings: 16 kA at 33 kV and 12.5 kA at 11 kV
- Selected 33 kV busbar conductor: ACSR Wolf
- Earthing grid dimensions: 48 m × 36 m
- Earth electrodes: 24 rods

The earthing assessment indicated acceptable touch and step voltage results under the assumed conditions. However, the ground potential rise requires further investigation.

The load growth assessment also indicates that the existing transformer capacity requires further planning from Year 4 under the assumed 5% annual growth rate.

## Project Files

- `Substation_Design_Calculations.xlsx` – Structured Excel workbook containing the design calculations.
- `drawings/` – Single-line diagram in DXF, PDF and PNG formats.
- `src/` – Python scripts for drawing generation, workbook creation and calculation verification.
- `docs/methodology.md` – Documentation of the calculation methodology.

## Verification and Limitations

I cross-checked the Excel calculation results using a separate Python verification script. The project reports 8 out of 10 checks passing, while two areas remain flagged for review: future transformer capacity and ground potential rise.

The calculations are based on assumed inputs and simplified engineering methods. Detailed cable sizing, battery/DC system design, lightning protection, civil design and other detailed engineering studies are outside the current scope.

The project is intended for learning and engineering practice. The assumptions, calculations, equipment ratings and safety-related results require further verification against actual network data, applicable standards and utility requirements before any real-world application.

## Future Scope

The project can be further developed by incorporating realistic network data, detailed cable sizing, earth-fault studies, additional protection studies and a Power BI dashboard for load growth analysis.

This project represents my effort to apply electrical engineering fundamentals, calculation-based design thinking, Python automation and technical documentation to a practical power distribution system.
