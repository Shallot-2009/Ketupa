##  [Notes]  Patch_antenna for Feed
##  V 0.0.2  Initial Release (Correlation with VC measurement data)
##  [Date]  Jul 14, 2025
##  [Source] From Asenjo.HB.L design model at Beijing.
##  [Author's Email]  3405802009@qq.com
##  [Copyright]  Copyright Asenjo.HB.L . All rights reserved.

![Teaser image](./assets/wifi_patch_antenna.bmp)
![Teaser image](./assets/Elliptical_Patch_Antenna.bmp)




## Introduction
This is a deep learning project applied to signal integrity and RF antenna design analysis.

## Installation
 
Clone this repository:

```
git clone https://github.com/Shallot-2009/Ketupa.git
cd ./Ketupa/
```


Install PyTorch and other dependencies:

```
conda create -y -n [ENV] python=3.8
conda activate [ENV]
```

```
### conda install -y pytorch=[>=1.6.0] torchvision cudatoolkit=[>=9.2] -c pytorch ###
### pip install torch==2.2.2 torchvision==0.17.2 torchaudio==2.2.2 --index-url https://download.pytorch.org/whl/cpu ###
```

```
pip install -r requirements.txt
python main.py
```


Ansys hfss models:

```
cd ./Ketupa/simulation/Ansys_HFSS/models/
 python  Wifi_patch_antenna.py
#python Edge_Fed_Rectangular_Patch_Antenna.py
#python Inset_Fed_Rectangular_Patch_Antenna.py
#python Planar_InvertedF_Antenna.py
#python Inset_Fed_Elliptical_Patch_Antenna.py
```

Matlab:

```
cd ./Ketupa/matlab/
run s_read_files.m
```

Output:
![Teaser image](./assets/wifi_patch_antenna_(s_matlab).bmp)
![Teaser image](./simulation/Ansys_HFSS/sim_results/Wifi_patch_antenna.bmp)

Matlab CNN Test:
![Teaser image](./assets/Matlab_CNN_test.bmp)
