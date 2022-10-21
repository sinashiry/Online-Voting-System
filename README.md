# Online Voting System
The Online Voting System is designed to conduct surveys in limited and decentralized places, which also considers security items during the voting.
The system works under the network and stores all the data received from the user in the internal network and on the central server.
The system consists of two parts:
Devices: An unlimited number of these devices can be added to the network and each of them can perform polling operations in a part of the environment. These units take a picture of a person’s face at the same time as conducting a survey, when entering an opinion, and save it on the server along with the data related to the opinion.
Software: This software is installed on the Windows operating system and allows the user to easily view the latest status of the system and devices, the amount and percentage of the voting (overall or by device), create a new voting, and view the voting in online mode.

The device has the ability to work long-term without interruption. The time and date of the system is set using the network clock and the main server, and it performs this operation in certain periods of time from the time it starts working.
If the device is offline in the network. It stores all input comments (+person’s faces) in the internal database and after establishing a stable connection with the server, it adapts all the data and deletes the internal data.

Additionally, A Windows based software is designed for manage and monitoring of devices and voting. software coded in python and used PyQt module. it connects to SQL Server for syncs results of all devices that is connects to network.
