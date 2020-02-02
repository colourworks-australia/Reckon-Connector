using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ReckonDesktop.Model
{

    public class Bills
    {
        public Accountingentry[] AccountingEntries { get; set; }
    }

    public class Accountingentry
    {
        public string id { get; set; }
        public string ContactID { get; set; }
        public string ContactName { get; set; }
        public string InvoiceNumber { get; set; }
        public string InvoiceDate { get; set; }
        public string AmountExTax { get; set; }
        public string AmountIncTax { get; set; }
        public string Account { get; set; }
        public string ItemType { get; set; }
        public string Description { get; set; }
        public string GlAccountBreakup { get; set; }
        public string Class { get; set; }
        public string JobNumber { get; set; }
    }

    public class AccountingLine
    {
        public string InvoiceNumber { get; set; }
        public string InvoiceDate { get; set; }
        public string AmountExTax { get; set; }
        public string AmountIncTax { get; set; }
        public string ItemType { get; set; }
        public string Description { get; set; }
        public string Qty { get; set; }
        public string ClassRef { get; set; }
    }

    public class AccountingContact
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        [Display(Name = "Contact Name")]
        public string Name { get; set; }
        [Display(Name = "ABN")]
        public string TaxNumber { get; set; }
        public int TenantId { get; set; }
        [Display(Name = "Contact Card No")]
        public string CardNumber { get; set; }
        public string ContactType { get; set; }
    }

    public class AccountingAccount
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public string AccountId { get; set; }
        public string AccountCode { get; set; }
        public string AccountDescription { get; set; }
        public int TenantId { get; set; }
    }

    public class AccountingJob
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public string JobId { get; set; }
        public string JobNumber { get; set; }
        public string JobName { get; set; }
        public int TenantId { get; set; }
        public string JobDescription { get; set; }
        public string ContactName { get; set; }
    }

    public class AccountingItemType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public string TypeName { get; set; }
        public int TenantId { get; set; }
    }
}
