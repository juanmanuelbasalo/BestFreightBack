using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Entities
{
    [Table("proveedor")]
    public class FreightProvider : BaseEntity
    {
        [Column("id")]
        [Key]
        public override int Id { get; set; }
        [Column("nombre")]
        public string Name { get; set; }
        [Column("contacto")]
        public string Contact { get; set; }
        [Column("email")]
        public string Email { get; set; }
        [Column("email2")]
        public string Email2 { get; set; }
        [Column("estado")]
        public int Status { get; set; }
        [Column("servicio")]
        public string Service { get; set; }
        [Column("pais")]
        public string Country { get; set; }

    }
}
