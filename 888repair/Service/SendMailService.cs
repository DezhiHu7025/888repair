using _888repair.Db;
using _888repair.Models.Director;
using _888repair.Models.Repair;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Service
{
    public class SendMailService
    {
        //送出报修单通知邮件
        public bool newMail(string AreaId, string KindId,string RepairId)
        {
            bool result = false;
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string findPerson = @"SELECT DISTINCT b.fullname,b.email
FROM [888_KsNorth].dbo.match a
    LEFT JOIN Common.dbo.kcis_account b
        ON a.charge_emp = b.EmpNo
WHERE (
          a.match_type = 'AreaMatch'
          AND a.area_id = @AreaId
      )
      OR
      (
          a.match_type = 'KindMatch'
          AND a.area_id = @KindId
      );";
                    var pList = db.Query<UserModel>(findPerson, new { AreaId, KindId }).ToList();
                    foreach (var p in pList)
                    {

                        EmailModel emailModel = new EmailModel();
                        emailModel.pid = "sys_flowengin";
                        emailModel.emailid = Convert.ToString(System.Guid.NewGuid());
                        emailModel.actiontype = "email";
                        emailModel.toaddr = p.email;
                        emailModel.toname = p.fullname;
                        emailModel.strSystem = "888报修系统(北)";
                        emailModel.subject = "888报修系统(北)--您收到一条报修单";
                        emailModel.body = string.Format(@"You have received a repair order:{0}. Please access to the website <a href='http://192.168.80.148/888repair_ksnorth/Home/Index?RPDetail-{0}' target='_blank' >[888Repair System(North)]</a>, and fill in it soon. Thank you! <br /><br />
                         您收到一条报修单：{0}。请您进入<a href='http://192.168.80.148/888repair_ksnorth/Home/Index?RPDetail-{0}' target='_blank' >[888报修系统(北)]</a> 尽快处理，谢谢！", RepairId);
                        string mailSql = string.Format(@"Insert into [Common].[dbo].[oa_emaillog](pid,emailid ,actiontype ,toaddr,toname ,fromaddr ,fromname,subject,body
                                  , attch , remark ,createdate ) 
                                   values(@pid, @emailid , @actiontype , @toaddr, @toname , 'automail@kcisec.com' , @strSystem, @subject, @body
                                  , @attch , @remark, getdate())");
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(mailSql, emailModel);
                        db.DoExtremeSpeedTransaction(trans);
                    }
                    
                }
                result = true;
                return result;
            }
            catch (Exception ex)
            {
                return result;
            }
        }

        //回应报修单通知邮件
        public bool replyMail(string RepairId, string Status)
        {
            bool result = false;
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string findPerson = @"SELECT a.*,
       b.fullname,
       b.email
FROM [888_KsNorth].dbo.record a
    LEFT JOIN Common.dbo.kcis_account b
        ON a.ResponseEmpno = b.EmpNo
WHERE a.repair_id = @RepairId ";
                    var pModel = db.Query<UserModel>(findPerson, new { RepairId }).FirstOrDefault();

                    string findStatus = @"SELECT b.StatusText
FROM [888_KsNorth].dbo.record a
    LEFT JOIN [888_KsNorth].dbo.state b
        ON a.Status = b.StatusValue
WHERE a.repair_id =@RepairId";
                    var StatusText = db.Query<string>(findStatus, new { RepairId }).FirstOrDefault();
                    EmailModel emailModel = new EmailModel();
                    emailModel.pid = "sys_flowengin";
                    emailModel.emailid = Convert.ToString(System.Guid.NewGuid());
                    emailModel.actiontype = "email";
                    emailModel.toaddr = pModel.email;
                    emailModel.toname = pModel.fullname;
                    emailModel.strSystem = "888报修系统(北)";
                    emailModel.subject = "888报修系统(北)--管理者回应了您的问题";
                    emailModel.body = string.Format(@"The repair report:{0} you sent has been responded to with a status of {1}.
                     Please access to the website <a href='http://192.168.80.148/888repair_ksnorth/Home/Index?PPDetail-{0}' target='_blank' >[888Repair System(North)]</a>, and fill in it soon. Thank you! <br /><br />
                     您送出的报修单:{0},已被回应，状态为{1}。请您进入<a href='http://192.168.80.148/888repair_ksnorth/Home/Index?PPDetail-{0}' target='_blank' >[888报修系统(北)]</a> 尽快处理，谢谢！", RepairId, StatusText);
                    string mailSql = string.Format(@"Insert into [Common].[dbo].[oa_emaillog](pid,emailid ,actiontype ,toaddr,toname ,fromaddr ,fromname,subject,body
                                  , attch , remark ,createdate ) 
                                   values(@pid, @emailid , @actiontype , @toaddr, @toname , 'automail@kcisec.com' , @strSystem, @subject, @body
                                  , @attch , @remark, getdate())");
                    Dictionary<string, object> trans = new Dictionary<string, object>();
                    trans.Add(mailSql, emailModel);
                    db.DoExtremeSpeedTransaction(trans);

                }
                result = true;
                return result;
            }
            catch (Exception ex)
            {
                return result;
            }
        }

        //转派报修单通知邮件
        public bool transferMail(string RepairId, string ChargeEmpno)
        {
            bool result = false;
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string findPerson = @"SELECT b.*
FROM  Common.dbo.kcis_account b
WHERE b.EmpNo = @ChargeEmpno ";
                    var pModel = db.Query<UserModel>(findPerson, new { ChargeEmpno }).FirstOrDefault();
                    EmailModel emailModel = new EmailModel();
                    emailModel.pid = "sys_flowengin";
                    emailModel.emailid = Convert.ToString(System.Guid.NewGuid());
                    emailModel.actiontype = "email";
                    emailModel.toaddr = pModel.email;
                    emailModel.toname = pModel.fullname;
                    emailModel.strSystem = "888报修系统(北)";
                    emailModel.subject = "888报修系统(北)--有一条报修单被转派给你";
                    emailModel.body = string.Format(@"A repair order:{0} has been transferred to you.Please access to the website <a href='http://192.168.80.148/888repair_ksnorth/Home/Index?RPDetail-{0}' onclick=queryDetail('{0}') target='_blank' >[888Repair System(North)]</a>, and fill in it soon. Thank you! <br /><br />
                      有一条报修单:{0}被转派给您。请您进入<a href='http://192.168.80.148/888repair_ksnorth/Home/Index?RPDetail-{0}' target='_blank' >[888报修系统(北)]</a> 尽快处理，谢谢！", RepairId);
                    string mailSql = string.Format(@"Insert into [Common].[dbo].[oa_emaillog](pid,emailid ,actiontype ,toaddr,toname ,fromaddr ,fromname,subject,body
                                  , attch , remark ,createdate ) 
                                   values(@pid, @emailid , @actiontype , @toaddr, @toname , 'automail@kcisec.com' , @strSystem, @subject, @body
                                  , @attch , @remark, getdate())");
                    Dictionary<string, object> trans = new Dictionary<string, object>();
                    trans.Add(mailSql, emailModel);
                    db.DoExtremeSpeedTransaction(trans);

                }
                result = true;
                return result;
            }
            catch (Exception ex)
            {
                return result;
            }
        }

        /// <summary>
        /// 驳回任务通知邮件
        /// </summary>
        /// <param name="RepairId"> 报修单单号</param>
        /// <param name="ChargeEmpno">上一任负责人 即驳回成功后的负责人</param>
        /// <param name="UpdateEmpNo">转派任务的操作人</param>
        /// <param name="EmpNo">驳回动作的操作人</param>
        /// <returns></returns>
        public bool rejectMail(string RepairId, string ChargeEmpno, string UpdateEmpNo,string FullName)
        {
            bool result = false;
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    //1.先判断当初转派报修单的人是不是当时的报修单负责人，是则只发一封提醒邮件，否则两人都要通知。
                    if (ChargeEmpno == UpdateEmpNo)
                    {
                        string findPerson = @"SELECT b.*
FROM  Common.dbo.kcis_account b
WHERE b.EmpNo = @ChargeEmpno ";
                        var pModel = db.Query<UserModel>(findPerson, new { ChargeEmpno }).FirstOrDefault();
                        EmailModel emailModel = new EmailModel();
                        emailModel.pid = "sys_flowengin";
                        emailModel.emailid = Convert.ToString(System.Guid.NewGuid());
                        emailModel.actiontype = "email";
                        emailModel.toaddr = pModel.email;
                        emailModel.toname = pModel.fullname;
                        emailModel.strSystem = "888报修系统(北)";
                        emailModel.subject = "888报修系统(北)--有一条报修单被驳回给您";
                        emailModel.body = string.Format(@"A repair order:{0} has been rejected by {1} for you.Please access to the website <a href='http://192.168.80.148/888repair_ksnorth/Home/Index?RPDetail-{0}' target='_blank' >[888Repair System(North)]</a>, and fill in it soon. Thank you! <br /><br />
                          有一条报修单:{0}被驳{1}回给您。请您进入<a href='http://192.168.80.148/888repair_ksnorth/Home/Index?RPDetail-{0}' target='_blank' >[888报修系统(北)]</a> 尽快处理，谢谢！", RepairId, FullName);
                        string mailSql = string.Format(@"Insert into [Common].[dbo].[oa_emaillog](pid,emailid ,actiontype ,toaddr,toname ,fromaddr ,fromname,subject,body
                                  , attch , remark ,createdate ) 
                                   values(@pid, @emailid , @actiontype , @toaddr, @toname , 'automail@kcisec.com' , @strSystem, @subject, @body
                                  , @attch , @remark, getdate())");
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(mailSql, emailModel);
                        db.DoExtremeSpeedTransaction(trans);
                    }
                    //给驳回的负责人发通知邮件，给转派任务的人发通知邮件
                    else
                    {
                        //驳回任务的操作人
                        string findPerson = @"SELECT b.*
FROM  Common.dbo.kcis_account b
WHERE b.EmpNo = @UpdateEmpNo ";
                        var pModel = db.Query<UserModel>(findPerson, new { UpdateEmpNo }).FirstOrDefault();

                        //上一任负责人
                        string findPerson2 = @"SELECT b.*
FROM  Common.dbo.kcis_account b
WHERE b.EmpNo = @ChargeEmpno ";
                        var pModel2 = db.Query<UserModel>(findPerson2, new { ChargeEmpno }).FirstOrDefault();
                        EmailModel emailModel = new EmailModel();
                        emailModel.pid = "sys_flowengin";
                        emailModel.emailid = Convert.ToString(System.Guid.NewGuid());
                        emailModel.actiontype = "email";
                        emailModel.toaddr = pModel.email;
                        emailModel.toname = pModel.fullname;
                        emailModel.strSystem = "888报修系统(北)";
                        emailModel.subject = "888报修系统(北)--您转派的报修单被驳回给上一任负责人";
                        emailModel.body = string.Format(@"A repair order:{0} has been transferred to you.Please access to the website <a href='http://192.168.80.148/888repair_ksnorth/Home/Index?APDetail-{0}' target='_blank' >[888Repair System(North)]</a>, and fill in it soon. Thank you! <br /><br />
                          您转派的报修单:{0}被{1}驳回给上一任负责人{2}。请您进入<a href='http://192.168.80.148/888repair_ksnorth/Home/Index?APDetail-{0}' target='_blank' >[888报修系统(北)]</a> 尽快处理，谢谢！", RepairId, FullName, pModel2.fullname);
                        string mailSql = string.Format(@"Insert into [Common].[dbo].[oa_emaillog](pid,emailid ,actiontype ,toaddr,toname ,fromaddr ,fromname,subject,body
                                  , attch , remark ,createdate ) 
                                   values(@pid, @emailid , @actiontype , @toaddr, @toname , 'automail@kcisec.com' , @strSystem, @subject, @body
                                  , @attch , @remark, getdate())");
                      

                        
                        EmailModel emailModel2 = new EmailModel();
                        emailModel2.pid = "sys_flowengin";
                        emailModel2.emailid = Convert.ToString(System.Guid.NewGuid());
                        emailModel2.actiontype = "email";
                        emailModel2.toaddr = pModel2.email;
                        emailModel2.toname = pModel2.fullname;
                        emailModel2.strSystem = "888报修系统(北)";
                        emailModel2.subject = "888报修系统(北)--有一条报修单被驳回给您";
                        emailModel2.body = string.Format(@"A repair order:{0} has been rejected by {1} for you.Please access to the website <a href='http://192.168.80.148/888repair_ksnorth/Home/Index?RPDetail-{0}' target='_blank' >[888Repair System(North)]</a>, and fill in it soon. Thank you! <br /><br />
                           有一条报修单:{0}被驳{1}回给您。 请您进入<a href='http://192.168.80.148/888repair_ksnorth/Home/Index?RPDetail-{0}' target='_blank' >[888报修系统(北)]</a> 尽快处理，谢谢！", RepairId, FullName);
                        string mailSql2 = string.Format(@"Insert into [Common].[dbo].[oa_emaillog](pid,emailid ,actiontype ,toaddr,toname ,fromaddr ,fromname,subject,body
                                  , attch , remark ,createdate ) 
                                   values(@pid, @emailid , @actiontype , @toaddr, @toname , 'automail@kcisec.com' , @strSystem, @subject, @body
                                  , @attch , @remark, getdate())");
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(mailSql, emailModel);
                        db.DoExtremeSpeedTransaction(trans);

                        Dictionary<string, object> trans2 = new Dictionary<string, object>();
                        trans2.Add(mailSql2, emailModel2);
                        db.DoExtremeSpeedTransaction(trans2);
                    }
                   
                }
                result = true;
                return result;
            }
            catch (Exception ex)
            {
                return result;
            }
        }

        public bool rejectEndMail(string RepairId)
        {
            bool result = false;
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string findPerson = @"SELECT a.*,
       b.fullname,
       b.email
FROM [888_KsNorth].dbo.record a
    LEFT JOIN Common.dbo.kcis_account b
        ON a.ResponseEmpno = b.EmpNo
WHERE a.repair_id = @RepairId ";
                    var pModel = db.Query<UserModel>(findPerson, new { RepairId }).FirstOrDefault();
                    EmailModel emailModel = new EmailModel();
                    emailModel.pid = "sys_flowengin";
                    emailModel.emailid = Convert.ToString(System.Guid.NewGuid());
                    emailModel.actiontype = "email";
                    emailModel.toaddr = pModel.email;
                    emailModel.toname = pModel.fullname;
                    emailModel.strSystem = "888报修系统(北)";
                    emailModel.subject = "888报修系统(北)--您的报修单被驳回";
                    emailModel.body = string.Format(@"The repair report:{0} you sent has been rejected.Please access to the website <a href='http://192.168.80.148/888repair_ksnorth/Home/Index?PPDetail-{0}' target='_blank' >[888Repair System(North)]</a>, and fill in it soon. Thank you! <br /><br />
                       您送出的报修单:{0},已被驳回。请您进入<a href='http://192.168.80.148/888repair_ksnorth/Home/Index?PPDetail-{0}' target='_blank' >[888报修系统(北)]</a> 尽快处理，谢谢！", RepairId);
                    string mailSql = string.Format(@"Insert into [Common].[dbo].[oa_emaillog](pid,emailid ,actiontype ,toaddr,toname ,fromaddr ,fromname,subject,body
                                  , attch , remark ,createdate ) 
                                   values(@pid, @emailid , @actiontype , @toaddr, @toname , 'automail@kcisec.com' , @strSystem, @subject, @body
                                  , @attch , @remark, getdate())");
                    Dictionary<string, object> trans = new Dictionary<string, object>();
                    trans.Add(mailSql, emailModel);
                    db.DoExtremeSpeedTransaction(trans);

                }
                result = true;
                return result;
            }
            catch (Exception ex)
            {
                return result;
            }
        }
    }
}