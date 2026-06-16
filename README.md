# Container Receiving Tool

A command-line tool that automates the presale-to-in-stock inventory 
transition process for container shipments in a steel door and window 
import/distribution operation.

## The Problem

Every time a new container arrived, an operations manager had to manually:
- Identify which presale products were included in the shipment
- Remove inventory counts from presale listings in Shopify
- Draft the presale product and transfer quantities to the in-stock version
- Do this for every product variant (size, finish, glass type, swing, hardware)

For a typical container of 80-100 units across 6-12 product lines with 
multiple variants each, this process took 2-3 hours per container arrival.

## The Solution

This tool reads a container manifest CSV and automatically generates:
- A structured receiving report grouped by product design
- Variant-level detail for every line item (size, finish, glass, swing)
- Add-on flagging with per-unit pricing
- Unit subtotals per product and grand total for the container
- A printable sign-off section for warehouse verification

What took 2-3 hours now takes seconds to generate.

## How to Use

1. Export your container manifest as a CSV with these columns:
   `Design Name, Quantity, Size, Finish, Glass Type, Top Shape, 
   Hardware, Add-Ons, Unit Price, Add-On Price, Swing`

2. Place the CSV in the same folder as the script and rename it 
   `container_manifest.csv`

3. Run the script:

4. The report prints to the terminal and saves to `receiving_log.txt`

## Sample Output

## Tech Stack

- Python 3
- Built with AI-assisted development (vibe coding)

## Background

Built as part of a portfolio project to demonstrate operational domain 
knowledge in steel door and window import/distribution, combined with 
AI-assisted tooling. The sample manifest uses fictional product names 
but reflects the real structure of container manifests from this industry.