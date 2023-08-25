# CasePilot

## Purpose
CasePilot is an internal tool designed to automate the paperwork associated with instances of academic dishonesty for St. Olaf College's Honor Council. Seamlessly woven into the Google Suite, it offers an intuitive user interface that facilitates easy utilization (as shown below).

<p align="center">
  <img src="https://raw.githubusercontent.com/radall1/CasePilot/main/info/pic1.png" />
</p>

## How-Tos? 
### How to use this tool?
For information on how to utilize the tool, please refer to the Honor Council's internal policies. 

### How to adjust CasePilot to accommodate changes in the Form Center?
Each question on the Form Center corresponds with a placeholder in the template(s) (one-to-many). CasePilot replaces the specific placeholder in the template(s) with the answer of the corresponding question on the Form Center. The correspondence is configured in the `case-log.gs` file as following:

 For example, the question "Permitted Resources/Collaboration" on the student-implicated side of the Form Center corresponds with the placeholder `{{resources}}` in the student-implicated Case Report template.


