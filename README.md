# viessmann-parameters-english

List of parameters for Viessmann heat pumps and such, translated in English.

This repository contains my own translation (from German) of lists of Viessmann parameters that I found online.
The reason for this repository to exist is to enable non-German-speakers to 
understand the register mappings of their Viessmann devices.

## Sources of Viessmann parameters

* [openv](https://github.com/openv/openv/wiki): OpenV project -- the wiki is a good source of information but it's entirely written in German
* [philippo66 processed Excel](https://github.com/philippoo66/optolink-splitter/wiki/310-Parameter-Addresses): This is possibly the best, organized source of Viessmann parameter addresses but again all parameters are written in German. 
* [philippo66 ViessData](https://github.com/philippoo66/ViessData21): This is another great website containing dumps of data coming from the VitoSoft software package.

## Translated data

Translation is mostly done and available here:

* [Parameter Addresses](Viessmann_Parameter_Adresses_EN_translated.xlsx): this is the direct German->English translation of [philippo66 processed Excel](https://github.com/philippoo66/optolink-splitter/wiki/310-Parameter-Addresses)

You will notice that the Excel contains a ton of sheets with model names like "Ecotronic", "CU401B_G", etc.
How do you know which one is yours? Keep reading the next section


## Identifying the model for your Viessmann Device

> What follows is a German->English translation of https://github.com/philippoo66/ViessData21/blob/master/README.md

First, the device identifier is read from address 0xF8 / 0x00F8. The 8 bytes have the following structure:

![graphic](https://user-images.githubusercontent.com/122479122/235326784-79c9e46b-fda9-4a53-b95e-04c68af3c47e.png)

In some cases, the `ecnsysDeviceIdentF0`—a single byte located at address 0xF0 / 0x00F0—is also required.

This information is then used to look up the corresponding entry in the [DataPoints list](https://github.com/philippoo66/ViessData21/blob/master/DataPoints.txt). For some devices, the first two bytes at F8 are sufficient—for example:

![graphic](https://user-images.githubusercontent.com/122479122/235326869-1710f9ce-dddb-42ec-ad36-c7da25bc8152.png)

For other devices, two additional bytes are required:

![graphic](https://user-images.githubusercontent.com/122479122/235326887-1ef0d12a-aeca-4fea-9ba3-07a54d06d455.png)

For a third group, the byte at F0 is also needed:

![graphic](https://user-images.githubusercontent.com/122479122/235326911-a20628f6-0313-4b93-80f5-0d81aca00369.png)

But BEWARE! Sometimes the `sysHardwareIndexIdent` '01' is merely a "placeholder" and can be interpreted as a "one-size-fits-all" value.

Example: On my Vitodens B3HB, the data at F8 is `20 CB 1F C9 00 00 01 14`, and at F0 it is `01`.

However, the `sysHardware/SoftwareIndexIdent` `1F C9` cannot be found anywhere. In fact, the hardware index is irrelevant here. It is a `VScotHO1_200_01`—or, in plain language: "software index 200 and up, Vitodens 300 project with HO2B, including ventilation control (as of 08/2016)." The 'from software index 200' designation is presumably found under `SoftwareIndexIdent` 0xC9 = 201.

For the purpose of identifying the corresponding data point list in [DP_Listen_2.zip](https://github.com/philippoo66/ViessData21/blob/master/DP_Listen_2.zip), the identifier following the colon on the right—in this example, 'VScotHO1_200_01'—is what matters. The associated data point list is `DP_VScotHO1_200_01.txt`.

ATTENTION: It appears that the F0 values ​​in `Devices.csv` are in decimal format rather than hexadecimal. For instance, `F0:0b` likely corresponds to 11–19.

## What to do next?

Check a project that allows you to connect your Viessmann device with MQTT like e.g.:

- [Optolink2MQTT](https://github.com/f18m/viessmann-optolink2mqtt): my own project, HomeAssistant friendly, modern take
- [Optolink Splitter](https://github.com/philippoo66/optolink-splitter): the original project that inspired this one
- [Optolink Bridge](https://github.com/kristian/optolink-bridge/): an approach based on the Optolink Splitter idea that requires a VitoConnect device and creates a man-in-the-middle bridge


## About Translator Software

See [translator README](./translator/README.md)
