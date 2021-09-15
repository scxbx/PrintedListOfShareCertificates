import os
from tkinter import filedialog

from msg_filler import MsgFiller
from msg_getter import VillageCommittee

if __name__ == '__main__':

    committee_path = filedialog.askdirectory(title='请选择村委会的文件夹')
    if committee_path != '':
        committee = VillageCommittee(committee_path)
        committee.sort_certificate()

        if not os.path.exists('结果'):
            os.mkdir('结果')
        c_path_list = os.path.split(committee_path)
        sample_path = r'模板\关于农村集体经济组织股权证书有关数据.xlsx'
        out_path = os.path.join(r'结果', c_path_list[1] + '.xlsx')
        filler = MsgFiller(committee.committee_list, sample_path, out_path)
        filler.fill()

        print('{} 填写完毕'.format(out_path))
        input('按任意键退出')

