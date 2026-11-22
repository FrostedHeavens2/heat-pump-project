# building-list

Per-building config files. Both `door-1-meters` and (eventually) `sizing-core`
read from here so a building's facts live in exactly one place.

Each file is an Excel sheet with two columns (`field`, `value`):

| field          | meaning                                                              |
| -------------- | -------------------------------------------------------------------- |
| name           | Short identifier (also used as filename).                            |
| occupancy      | People served.                                                       |
| floors         | Number of floors (optional, informational).                          |
| zip_code       | Building ZIP (ecosizer climate zone lookup).                         |
| supply_temp_f  | Hot water supply temperature out of fixtures (°F).                   |
| storage_temp_f | Storage tank setpoint (°F).                                          |
| inlet_temp_f   | Cold-water inlet temperature assumption (°F).                        |
| schematic      | ecosizer schematic: parallel_loop, swing_tank, single_pass_rtp, etc. |
| building_type  | ecosizer building type, e.g. multi_family.                           |
| notes          | Free text.                                                           |

Use `_template.xlsx` as a starting point. Name new files after the building,
e.g. `my_building.xlsx`.
