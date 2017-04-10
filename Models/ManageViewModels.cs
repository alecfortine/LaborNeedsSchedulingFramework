using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNet.Identity;
using Microsoft.Owin.Security;
using System.Data;
using System.Data.Entity;
using System;

namespace LaborNeedsSchedulingFramework.Models
{

    /// <summary>
    /// 
    /// </summary>
    public class DataInput
    {
        public DataTable managerTable { get; set; }
        public double payRoll { get; set; }
        public int minEmp { get; set; }
        public int maxEmp { get; set; }

        //weighting variables
        public double week1Weight { get; set; }
        public double week2Weight { get; set; }
        public double week3Weight { get; set; }
        public double week4Weight { get; set; }
        public double week5Weight { get; set; }
        public double week6Weight { get; set; }

        #region date times

        public string sunday0 { get; set; }
        public string monday0 { get; set; }
        public string tuesday0 { get; set; }
        public string wednesday0 { get; set; }
        public string thursday0 { get; set; }
        public string friday0 { get; set; }
        public string saturday0 { get; set; }


        public string sunday6 { get; set; }
        public string sunday5 { get; set; }
        public string sunday4 { get; set; }
        public string sunday3 { get; set; }
        public string sunday2 { get; set; }
        public string sunday1 { get; set; }

        public string monday6 { get; set; }
        public string monday5 { get; set; }
        public string monday4 { get; set; }
        public string monday3 { get; set; }
        public string monday2 { get; set; }
        public string monday1 { get; set; }

        public string tuesday6 { get; set; }
        public string tuesday5 { get; set; }
        public string tuesday4 { get; set; }
        public string tuesday3 { get; set; }
        public string tuesday2 { get; set; }
        public string tuesday1 { get; set; }

        public string wednesday6 { get; set; }
        public string wednesday5 { get; set; }
        public string wednesday4 { get; set; }
        public string wednesday3 { get; set; }
        public string wednesday2 { get; set; }
        public string wednesday1 { get; set; }

        public string thursday6 { get; set; }
        public string thursday5 { get; set; }
        public string thursday4 { get; set; }
        public string thursday3 { get; set; }
        public string thursday2 { get; set; }
        public string thursday1 { get; set; }

        public string friday6 { get; set; }
        public string friday5 { get; set; }
        public string friday4 { get; set; }
        public string friday3 { get; set; }
        public string friday2 { get; set; }
        public string friday1 { get; set; }

        public string saturday6 { get; set; }
        public string saturday5 { get; set; }
        public string saturday4 { get; set; }
        public string saturday3 { get; set; }
        public string saturday2 { get; set; }
        public string saturday1 { get; set; }
        #endregion

        #region date bools
        public bool sunday6check { get; set; }
        public bool sunday5check { get; set; }
        public bool sunday4check { get; set; }
        public bool sunday3check { get; set; }
        public bool sunday2check { get; set; }
        public bool sunday1check { get; set; }

        public bool monday6check { get; set; }
        public bool monday5check { get; set; }
        public bool monday4check { get; set; }
        public bool monday3check { get; set; }
        public bool monday2check { get; set; }
        public bool monday1check { get; set; }

        public bool tuesday6check { get; set; }
        public bool tuesday5check { get; set; }
        public bool tuesday4check { get; set; }
        public bool tuesday3check { get; set; }
        public bool tuesday2check { get; set; }
        public bool tuesday1check { get; set; }

        public bool wednesday6check { get; set; }
        public bool wednesday5check { get; set; }
        public bool wednesday4check { get; set; }
        public bool wednesday3check { get; set; }
        public bool wednesday2check { get; set; }
        public bool wednesday1check { get; set; }

        public bool thursday6check { get; set; }
        public bool thursday5check { get; set; }
        public bool thursday4check { get; set; }
        public bool thursday3check { get; set; }
        public bool thursday2check { get; set; }
        public bool thursday1check { get; set; }

        public bool friday6check { get; set; }
        public bool friday5check { get; set; }
        public bool friday4check { get; set; }
        public bool friday3check { get; set; }
        public bool friday2check { get; set; }
        public bool friday1check { get; set; }

        public bool saturday6check { get; set; }
        public bool saturday5check { get; set; }
        public bool saturday4check { get; set; }
        public bool saturday3check { get; set; }
        public bool saturday2check { get; set; }
        public bool saturday1check { get; set; }
        #endregion

        // power hours 
    }

    public class IndexViewModel
    {
        public bool HasPassword { get; set; }
        public IList<UserLoginInfo> Logins { get; set; }
        public string PhoneNumber { get; set; }
        public bool TwoFactor { get; set; }
        public bool BrowserRemembered { get; set; }
    }

    public class ManageLoginsViewModel
    {
        public IList<UserLoginInfo> CurrentLogins { get; set; }
        public IList<AuthenticationDescription> OtherLogins { get; set; }
    }

    public class FactorViewModel
    {
        public string Purpose { get; set; }
    }

    public class SetPasswordViewModel
    {
        [Required]
        [StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "New password")]
        public string NewPassword { get; set; }

        [DataType(DataType.Password)]
        [Display(Name = "Confirm new password")]
        [Compare("NewPassword", ErrorMessage = "The new password and confirmation password do not match.")]
        public string ConfirmPassword { get; set; }
    }

    public class ChangePasswordViewModel
    {
        [Required]
        [DataType(DataType.Password)]
        [Display(Name = "Current password")]
        public string OldPassword { get; set; }

        [Required]
        [StringLength(100, ErrorMessage = "The {0} must be at least {2} characters long.", MinimumLength = 6)]
        [DataType(DataType.Password)]
        [Display(Name = "New password")]
        public string NewPassword { get; set; }

        [DataType(DataType.Password)]
        [Display(Name = "Confirm new password")]
        [Compare("NewPassword", ErrorMessage = "The new password and confirmation password do not match.")]
        public string ConfirmPassword { get; set; }
    }

    public class AddPhoneNumberViewModel
    {
        [Required]
        [Phone]
        [Display(Name = "Phone Number")]
        public string Number { get; set; }
    }

    public class VerifyPhoneNumberViewModel
    {
        [Required]
        [Display(Name = "Code")]
        public string Code { get; set; }

        [Required]
        [Phone]
        [Display(Name = "Phone Number")]
        public string PhoneNumber { get; set; }
    }

    public class ConfigureTwoFactorViewModel
    {
        public string SelectedProvider { get; set; }
        public ICollection<System.Web.Mvc.SelectListItem> Providers { get; set; }
    }
}