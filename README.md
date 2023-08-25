# CasePilot

## Purpose
CasePilot is an internal tool designed to automate the paperwork associated with instances of academic dishonesty for St. Olaf College's Honor Council. Seamlessly woven into the Google Suite, it offers an intuitive user interface that facilitates easy utilization (as shown below).

<p align="center">
  <img src="https://raw.githubusercontent.com/radall1/CasePilot/main/info/pic1.png" />
</p>

## Notes
1. For information on how to utilize the tool, please refer to the Honor Council's internal policies. 

2. Each question on the Form Center corresponds with a placeholder in the template(s) (one-to-many). CasePilot replaces the specific placeholder in the template(s) with the answer of the corresponding question on the Form Center. The correspondence is configured in the `case-log.gs` file as follows:

```javascript
// Define placeholders for student and faculty cases
const PLACEHOLDERS_STUDENT = {
  'case number': 0,
...
  'resources': 35
};
```
In this case, for example, the placeholder `{{resources}}` in the student-implicated Case Report template corresponds with column number 35 (start counting from zero, not one) in the Case Log spreadsheet which is the question "Permitted Resources/Collaboration" on the student-implicated side of the Form Center. Therefore, if you change the ordering of the question (e.g. add or remove a question) to the Form Center, you will need to check that the linking is still correct. However, changing the text of the questions will not impact CasePilot.

3. The templates IDs operated on by this version of CasePilot are defined on the top lines of `case-log.gs` as follows:
```javascript
// Define named constants for template IDs and destination folder ID
const TEMPLATE_STUDENT_0 = '1YZeAV9U9QGehjz89AYtw2I_jhsq_6xgBO32ASH-k2do';      // URL for Case Report
const TEMPLATE_STUDENT_1 = '1XVDjJXlKSvu2UbuWFNk2__M0LVT4a3yi2wrlmq44nzA';      // URL for Notification Memo
```
The template ID can be derived from the link of the Google Doc. If the link is "docs.google.com/document/d/1YZeAV9U29QGehjz89AYtw2I_jhsq_6xgBO32ASH-k2do/edit" then the template ID is "1YZeAV9U29QGehjz89AYtw2I_jhsq_6xgBO32ASH-k2do".
