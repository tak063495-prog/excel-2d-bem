"""One entry point: python run_bem.py model.json --out output"""
import argparse,json
from pathlib import Path
from threadpoolctl import threadpool_limits
from elasticbem.io import save

def main():
    parser=argparse.ArgumentParser(description=__doc__); parser.add_argument('model'); parser.add_argument('--out',default='output'); parser.add_argument('--threads',type=int,default=1); args=parser.parse_args()
    if args.threads<1: parser.error('--threads must be positive.')
    model=json.loads(Path(args.model).read_text(encoding='utf-8-sig'))
    with threadpool_limits(limits=args.threads): return save(model,args.out)
if __name__=='__main__': raise SystemExit(main())
