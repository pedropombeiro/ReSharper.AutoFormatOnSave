// --------------------------------------------------------------------------------------------------------------------
// <copyright company="Pedro Pombeiro" file="GuidList.cs">
//   2012 Pedro Pombeiro
// </copyright>
// --------------------------------------------------------------------------------------------------------------------
namespace ReSharper.AutoFormatOnSave
{
    using System;

    internal static class GuidList
    {
        #region Constants

        public const string guidReSharper_AutoFormatOnSaveCmdSetString = "ed56bc7b-c00c-4949-ae12-ca9f33c962dd";

        public const string guidReSharper_AutoFormatOnSavePkgString = "2c65fba8-3ecb-4004-8ee3-6d9e62cd987a";

        #endregion

        #region Static Fields

        public static readonly Guid guidReSharper_AutoFormatOnSaveCmdSet = new Guid(guidReSharper_AutoFormatOnSaveCmdSetString);

        #endregion
    };
}