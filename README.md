```markdown
# Config Generator

This repository contains a Python script to generate XML configuration files from an Excel file with sheets `NP` and `SP`. The script processes one or more SITE names (provided as comma-separated values) and creates separate XML configuration, log, and var files for each SITE.

## Author

- **Vijayakumar R**
- **CFX-5000 Team**
- **Email:** vijaya.r.ext@nokia.com

## Files

- ```python
  e2503_orb_config_generator_xml.py
  ```
  This is the main script that:
  - Reads configuration parameters from the Excel file.
  - Updates an XML template based on the provided data.
  - Updates a deploy group with default parameters and prompts for `TrafficFileName`.
  - Writes the final XML configuration.
  - Generates accompanying log and variable files.

- ```text
  requirements.txt
  ```
  Lists required Python packages: `openpyxl` and `lxml`.

## Usage

1. **Prepare your Excel file:**  
   The Excel file must have sheets named `NP` and `SP`. Each sheet must include:
   - **NE Parameter Name**: The key for each parameter.
   - **Resource Sub-Type**:  
     - For non-ZTS groups, use the format `cfx^GroupName` (e.g. `cfx^OfflineParameter`).
     - For ZTS groups, use the format `cfx^ZTS^groupname` (e.g. `cfx^ZTS^ocsp_config`).
   - Additional columns for various SITE names. The values under these columns are used for configuration.

2. **Prepare your XML Template file:**  
   The XML template must include:
   - A container element `<cfx:cfx>` for non-ZTS groups.
   - A container element `<zts_cm:ZTS>` for ZTS groups.

3. **Run the script:**  
   Execute the script by running:
   ```bash
   python e2503_orb_config_generator_xml.py
   ```
   When prompted:
   - Provide the full path to the Excel file.
   - Provide the full path to the XML Template file.
   - Enter comma-separated SITE column names (e.g., `site1, site2`).
   - Enter the output directory where the configuration, log, and var files should be saved.
   
   The script will generate for each SITE:
   - `<SITE>_config.xml` — the generated XML configuration.
   - `<SITE>_config.log` — a log file documenting processing details.
   - `<SITE>_config.vars` — a file containing deploy group parameters.

4. **Deploy Group & Manual Changes:**  
   After the XML is generated, please note the following manual updates:
   - **ZTS Section:**  
     After the closing tag `</zts_common:zts-features>`, remove all lines where the values are `n/a`.
   - **Deploy Group:**  
     The following deploy group parameters are hardcoded in the XML:
     - `<deploy:ZtsLcmUsername>` is hardcoded as `zts1user`.
     - `<deploy:ZtsLcmPassword>` is hardcoded as `********`.
     - `<deploy:CncsUMAdminPassword>` is hardcoded as `none`.
     
     In the GS Latest configuration, these values are not present. You must add these manually if needed.
   - **TrafficFileName:**  
     `<deploy:TrafficFileName>` is generated based on user input. Replace the GS default Traffic file name with the actual traffic file name.
   - **Parameter Modification:**  
     Change:
     ```xml
     <pm-nbi-prometheus-enabled>false</pm-nbi-prometheus-enabled>
     ```
     to:
     ```xml
     <zts_efs:pm-nbi-rtpm-prometheus-enabled>false</zts_efs:pm-nbi-rtpm-prometheus-enabled>
     ```
     (Our script corrects this automatically using the parameter corrections mapping.)

## Requirements

Install the required Python packages using:
```bash
pip install -r requirements.txt
```

## Note

Hardcoded in XML output from the script:
- For the deploy group, the values for `<deploy:ZtsLcmUsername>`, `<deploy:ZtsLcmPassword>`, and `<deploy:CncsUMAdminPassword>` are set to `zts1user`, `********`, and `none` respectively.  
- `<deploy:TrafficFileName>` is derived from user input.
- After the XML is generated in the ZTS section, manually remove any lines below `</zts_common:zts-features>` that contain values of `n/a`.
- The parameter `<pm-nbi-prometheus-enabled>` is automatically corrected to `<zts_efs:pm-nbi-rtpm-prometheus-enabled>`.

Feel free to adjust the script and mappings as needed.
```