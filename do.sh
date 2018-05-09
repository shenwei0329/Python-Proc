#!/bin/sh
#

sh << EOF

cd /home/shenwei/python-proc

python DataHandler/jira_class_epic.py GZ story
python DataHandler/jira_class_epic.py GZ task

python DataHandler/jira_class_epic.py JX task

python DataHandler/jira_class_epic.py RDM epic
python DataHandler/jira_class_epic.py TESTCENTER task

python DataHandler/jira_class_epic.py CPSJ task

python DataHandler/jira_class_epic.py ROOOT story
python DataHandler/jira_class_epic.py ROOOT task

python DataHandler/jira_class_epic.py HUBBLE epic
python DataHandler/jira_class_epic.py HUBBLE task

python DataHandler/jira_class_epic.py FAST epic
python DataHandler/jira_class_epic.py FAST task

EOF
