"""Reproducible chart from generated base forecast CSV."""
import argparse
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd


def plot(path: str, output: str) -> None:
    frame=pd.read_csv(path)
    fig,ax=plt.subplots(figsize=(9,4.3),layout="constrained")
    fig.patch.set_facecolor("#f6f8fb");ax.set_facecolor("#f6f8fb")
    bars=ax.bar(frame.year,frame.fcff,color="#286f78",width=.55)
    ax.bar_label(bars,fmt="%.1f",padding=4)
    ax.set(title="Illustrative company · Base unlevered free cash flow",xlabel="Forecast year",ylabel="USD millions")
    ax.spines[["top","right"]].set_visible(False)
    ax.set_xticks(frame.year);ax.set_ylim(min(0,frame.fcff.min()*1.2),frame.fcff.max()*1.2)
    fig.text(.01,.01,"Fictional training data. No real security or investment recommendation.",fontsize=8,color="#536577")
    Path(output).parent.mkdir(parents=True,exist_ok=True);fig.savefig(output,dpi=170);plt.close(fig)


if __name__=="__main__":
    parser=argparse.ArgumentParser();parser.add_argument("forecast");parser.add_argument("output")
    args=parser.parse_args();plot(args.forecast,args.output)
