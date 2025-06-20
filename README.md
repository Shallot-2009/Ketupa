##  [Notes]  wifi_patch_antenna for 2.4GHz
##                 V 0.0.1  Initial Release (Correlation with VC measurement data)
##  [Date]  Jun 19, 2025
##  [Source] From Asenjo.HB.L design model at Beijing.
##  [Author's Email]  3405802009@qq.com
##  [Copyright]  Copyright Asenjo.HB.L . All rights reserved.

![Teaser image](./assets/wifi_patch_antenna.bmp)




## Introduction
This is a deep learning project applied to signal integrity and RF analysis.

## Installation
 
Clone this repository:

```
git clone https://github.com/Shallot-2009/Ketupa.git
cd ./Ketupa_demo/
```



Install PyTorch and other dependencies:

```
conda create -y -n [ENV] python=3.8
conda activate [ENV]
###conda install -y pytorch=[>=1.6.0] torchvision cudatoolkit=[>=9.2] -c pytorch##
pip install torch==2.2.2 torchvision==0.17.2 torchaudio==2.2.2 --index-url https://download.pytorch.org/whl/cpu
cd ./Ketupa_demo/
pip install -r requirements.txt
cd ./Ketupa_demo/HFSS/
python wifi_patch_antenna.py
```
