# -*- coding: UTF-8 -*-
import itchat
from itchat.content import *
import ExcelRW
import time
import threading
import re
#from queue import Queue


""" macros definition"""
JUDGE_DELAY = 180    # the delay time to judge if all message are sent for that day,unit: seconds
TIME_SCORE_DIV = 10000   # the division between voice time and score or sentence
#GROUP_NAME = u'I  ❤️ B.O.A.T.'
GROUP_NAME = u'Test'

# names can not be changed, must match the "Group Alias" in BOAT group
#Member_names = {'sherry':1, 'Chris':1, 'Joanna':1,
#               'Isabel':2, 'Leo':2, 'Lincoln':2,
#                'Demi':3, 'Doris':3, 'Mike':3}
                # question: what if a participant has no group alias?
Member_names = {'Chris':1, 'Chrisnes':2, 'lee':1, 'noone':3}


Names_list = list(Member_names) # get a list of all members
Participants = []   # list the basic information of each member
msg_list = []       # list the current existing message
thread_list = []    # for test only, see what task in on the list
Current_day_flag = time.localtime(time.time()).tm_mday


class Gamers(object):
    # members definition
    def __init__(self, index, name, group_index):
        self.__index = index  # index number of a member, from 0 to 8
        self.__name = name    # name of that member, is the group alies in BOAT
        self.__group = group_index    # group that member belongs to
        self.__today_score = 0        # the score the member has got today, cleared every day, type: int

    def add_score(self, total_voice_len):
        """add score on a member.

        :param total_voice_len: the total voice length(time), unit: micro seconds
        :type total_voice_len: int

        """
        temp_score = total_voice_len // TIME_SCORE_DIV
        self.__today_score += temp_score  # score validity will be checked in function save_score()
        self.save_score()

    def reset_everyday_task(self):
        """reset all data in a day.

        """
        self.__today_score = 0

    def save_score(self):
        """save the score of today in excel sheet.

        """
        ExcelRW.excel_save(index=self.__index, score=self.__today_score)

    def get_member_index(self):
        return self.__index

    def get_member_name(self):
        return self.__name

    def get_group_index(self):
        return self.__group


class MsgQueue(object):
    def __init__(self, participant_index, member_name):
        self.__index = participant_index      # index of the member, it corresponds to the index in Gamers class
        self.__member_name = member_name
        self.start_time = None  # if it is None, the member has not send any voice yet
        self.msg_count = 0  # number of message the member has sent in a period (dynamic 10 minutes)
        self.total_voice_len = 0    # total valid voice time
        self.msg_queue = {}     # a temporary dict to save all the voice message that member has sent in a period of time

    def get_index(self):
        return self.__index

    def get_name(self):
        return self.__member_name

    def add_msg(self, msg_id, voice_len, time_stamp):  # if receive is a voice message
        """add a message.

        :param msg_id: the id number of the voice message get from itchat.msg. type: int
        :param voice_len: the voice len(microseconds) of a voice message    Type: int
        :param time_stamp: the time stamp of the last voice message

        """
        self.start_time = time_stamp
        self.msg_count += 1
        temp = {msg_id:voice_len}
        self.msg_queue.update(temp) # save this message in a buffer

    def del_msg(self, msg_id, time_stamp):  # if recalled a voice message
        """delete a message.

        :param msg_id: the id number of the voice message get from itchat.msg. type: int
        :param time_stamp: time stamp of the notice

        """
        if self.msg_count >= 1: # to avoid mistake that a member has no message but want to delete
            self.msg_count -=1
            if msg_id in self.msg_queue:    # must check if msg_id is valid or not
                del self.msg_queue[msg_id]
                self.start_time = time_stamp

    def clear_msg(self):
        """clear all the message buffer of a member

        """
        self.start_time = None
        self.msg_count = 0
        self.total_voice_len = 0
        self.msg_queue.clear()

    def cal_total_voice_len(self):
        """calculate score according to the total valid voice message a member has sent in a day

        """
        temp_voice_len_list = self.msg_queue.values()   # return a list with all the voice length
        for each_len in temp_voice_len_list:
            self.total_voice_len += each_len        # tantalize all the voice message length


def init():
    """initialize the variables and parameters

    """
    for each_member in Names_list:
        index_num = Names_list.index(each_member)
        group_index = Member_names[each_member]
        temp = Gamers(index=index_num, name=each_member, group_index=group_index)
        Participants.append(temp)       # create a list of members
        temp = MsgQueue(participant_index=index_num, member_name=each_member)
        msg_list.append(temp)          # create a list of message queue

        ExcelRW.excel_init(members_dict=Member_names)


@itchat.msg_register([VOICE], isGroupChat=True, isFriendChat=False, isMpChat=False)
def message_check(msg):
    """check message, it is called when a voice message is sent from BOAT group

    :param msg: the msg get from itchat module, include all the information related to the message in Wechat

    """
    if msg.user.NickName == GROUP_NAME:      # if the message is from group BOAT
        if msg.ActualNickName in Names_list:  # if sender is in the game member list
            index = Names_list.index(msg.ActualNickName)
        elif msg.ActualNickName == '':  # Chris     this code can be better
            index = Names_list.index('Chris')
        else:
            return
        cur_time = time.time()
        voice_len = msg.VoiceLength
        msg_list[index].add_msg(msg_id=msg.MsgId, voice_len=voice_len, time_stamp=cur_time)
    else:
        pass


@itchat.msg_register([NOTE], isGroupChat=True, isFriendChat=False, isMpChat=False)
def message_recall(msg):
    """deal with the recalled message

    :param msg: the msg get from itchat module, include all the information related to the message in Wechat

    """
    if re.search('<sysmsg type="revokemsg">', msg.content) is None: # if it is not a message recall note
        return
    if msg.user.NickName == GROUP_NAME:      # if the message is from group BOAT
        if msg.ActualNickName in Names_list:
            index = Names_list.index(msg.ActualNickName)
        elif msg.ActualNickName == '':  # Chris
            index = Names_list.index('Chris')
        else:
            return
        start = re.search('<msgid>', msg.content).span()
        end = re.search('</msgid>', msg.content).span()
        if start is not None and end is not None:
            old_msg_id = msg.content[start[1]:end[0]]    # find the old message id in msg.content
            time_stamp = time.time()
            msg_list[index].del_msg(msg_id=old_msg_id, time_stamp=time_stamp)


def msg_que_check():
    """Check message queue and calculate if it is valid or not

    """
    while True:
        for each_msg in msg_list:
            if each_msg.start_time is not None:
                accumulating_time = time.time() - each_msg.start_time
                if accumulating_time > JUDGE_DELAY:
                    each_msg.cal_total_voice_len()
                    Participants[each_msg.get_index()].add_score(total_voice_len=each_msg.total_voice_len)
                    send_wechat_msg(member_index=each_msg.get_index()) #send message to the member to inform the result
                    each_msg.clear_msg()

        global Current_day_flag
        if time.localtime(time.time()).tm_hour is 0 and \
                Current_day_flag != time.localtime(time.time()).tm_mday:    # if it is another day, (after 00:00)
            for each_msg in msg_list:
                each_msg.clear_msg()
            for each_member in Participants:
                each_member.reset_everyday_task()
            Current_day_flag = time.localtime(time.time()).tm_mday


def send_wechat_msg(member_index):
    """send wechat message

    :param member_index: the index number of that member that I am going to mention

    """
    scores = ExcelRW.read_score(member_index=member_index, total_members=len(Member_names))
    member_name = Participants[member_index].get_member_name()
    member_score = scores[0]
    group_number = Participants[member_index].get_group_index()
    group_score = scores[1]

    # can I @ that member?
    message_content = "[AUTO REPLY]: @%s: Score points get today is %d. \nTotal score of group %s is %d" \
                      % (member_name, member_score, group_number, group_score)
    itchat.send_msg(message_content)


def main_loop():
    t1 = threading.Thread(target=itchat.run)    # register itchat run in thread
    thread_list.append(t1)
    t2 = threading.Thread(target=msg_que_check)     # register msg_que_check in thread
    thread_list.append(t2)

    for t in thread_list:
        t.start()
    for t in thread_list:
        t.join()


if __name__ == '__main__':
    init()
    #itchat.auto_login()
    main_loop()