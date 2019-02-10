using System;
using System.Linq;
using System.Data.Entity;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace FIS
{
    public class Izdel
    {
        [Key]
        public System.Int64 Id { get; set; }
        public System.String Name { get; set; }
        public System.Decimal Price { get; set; }

    }

    public class Link
    {
        [Key]
        public System.Int64 Id { get; set; }
        public System.Int64 IzdelUp { get; set; }
        public System.Int32 Kol { get; set; } /*Количество текущих изделий, входящих в вышестоящее*/
        [ForeignKey("Izdel")]
        public System.Int64 IzdelId { get; set; }
        public Izdel Izdel { get; set; }
    }

    class Product
    {
        public System.Int64 Id { get; set; }
        public System.String Name { get; set; }
        public System.Decimal Price { get; set; }
        public System.Int64 Parent { get; set; }
        public System.Int32 Quantity { get; set; }
        public System.Boolean Label { get; set; }
        public System.Int32 ChildQuantity { get; set; }
        public List<Product> children = new List<Product>();
        public Product(long id, string name, decimal price, long parent, int quantity)
        {
            Id = id;
            Name = name;
            Price = price;
            Parent = parent;
            Quantity = quantity;
        }
        public Product() { }
    }

    class FIScontext : DbContext
    {
        public FIScontext()
            : base("DbConnection")
        { }
        
        public DbSet<Izdel> Izdels { get; set; }
        public DbSet<Link> Links { get; set; }

    }

    public static class Database
    {
        public static void Add(Izdel izdel, Link link)
        {
            using (FIScontext db = new FIScontext())
            {
                if(izdel != null && link != null)
                {
                    db.Izdels.Add(izdel);
                    db.Links.Add(link);
                }
                db.SaveChanges();
            }
        }

        public static void AddIzdel(Izdel izdel)
        {
            using (FIScontext db = new FIScontext())
            {
                if (izdel != null)
                {
                    db.Izdels.Add(izdel);
                    db.SaveChanges();
                }
            }
        }
        public static void DeleteIzdel(Izdel izdel)
        {
            using (FIScontext db = new FIScontext())
            {
                var _izdel = db.Izdels.FirstOrDefault(I => I.Id == izdel.Id);
                if (_izdel != null)
                {
                    db.Izdels.Remove(_izdel);
                    db.Links.Remove(db.Links.FirstOrDefault(L=>L.IzdelId==_izdel.Id));
                    db.SaveChanges();
                }
            }
        }
        public static void DeleteIzdel()
        {
            using (FIScontext db = new FIScontext())
            {
                foreach (var izdel in db.Izdels.ToList())
                {
                    db.Izdels.Remove(izdel);
                    var _link = db.Links.FirstOrDefault(L => L.IzdelId == izdel.Id);
                    if (_link != null)  db.Links.Remove(_link);
                }
                db.SaveChanges();
            }
        }
        public static void AddLink(Link link)
        {
            using (FIScontext db = new FIScontext())
            {
                if (link != null)
                {
                    db.Links.Add(link);
                    db.SaveChanges();
                }
            }
        }
        public static void DeleteLink(Link link)
        {
            using (FIScontext db = new FIScontext())
            {
                var _link = db.Links.FirstOrDefault(l => l.Id == link.Id);
                if (_link != null)
                {
                    db.Links.Remove(_link);
                    db.SaveChanges();
                }
            }
        }
        public static void DeleteLink()
        {
            using (FIScontext db = new FIScontext())
            {
                foreach (var link in db.Links.ToList())
                    db.Links.Remove(link);
                db.SaveChanges();
            }
        }

        //Вывод изделий
        public static void PrintIzdel()
        {
            using (FIScontext db = new FIScontext())
            {
                if (db.Izdels.Count() != 0)
                {
                    foreach (var izdel in db.Izdels.ToList())
                    {
                        Console.Write("{0} {1} {2} ", izdel.Id, izdel.Name, izdel.Price);
                        var link = db.Links.FirstOrDefault(L => L.IzdelId == izdel.Id);
                        if (link != null)
                        {
                            Console.WriteLine("{0} {1}", link.IzdelUp, link.Kol);
                        }
                        else
                        {
                            Console.WriteLine("Нет link");
                        }
                    }
                }
                else
                    Console.WriteLine("Нет изделий");
            }
        }

        public static void initDatabase()
        {
            Random rnd = new Random();
            Izdel[] izdels = new Izdel[10];
            Link[] links = new Link[10];
            for (int i = 0; i < 10; i++)
                izdels[i] = new Izdel {Id = (i+1), Name = "Изделие"+(i+1), Price = 10 + rnd.Next(100) / 10 * 10 };
            links[0] = new Link { IzdelUp = -1, IzdelId = izdels[0].Id, Kol = 1 };
            links[1] = new Link { IzdelUp = izdels[0].Id, IzdelId = izdels[1].Id, Kol = 1 + rnd.Next(10)};
            links[2] = new Link { IzdelUp = izdels[0].Id, IzdelId = izdels[2].Id, Kol = 1 + rnd.Next(10) };
            links[3] = new Link { IzdelUp = izdels[2].Id, IzdelId = izdels[3].Id, Kol = 1 + rnd.Next(10) };
            links[4] = new Link { IzdelUp = izdels[0].Id, IzdelId = izdels[4].Id, Kol = 1 + rnd.Next(10) };
            links[5] = new Link { IzdelUp = izdels[4].Id, IzdelId = izdels[5].Id, Kol = 1 + rnd.Next(10) };
            links[6] = new Link { IzdelUp = izdels[4].Id, IzdelId = izdels[6].Id, Kol = 1 + rnd.Next(10) };
            links[7] = new Link { IzdelUp = -1, IzdelId = izdels[7].Id, Kol = 1 };
            links[8] = new Link { IzdelUp = izdels[7].Id, IzdelId = izdels[8].Id, Kol = 1 + rnd.Next(10) };
            links[9] = new Link { IzdelUp = izdels[7].Id, IzdelId = izdels[9].Id, Kol = 1 + rnd.Next(10) };
            for (int i = 0; i < 10; i++)
                {
                    Add(izdels[i], links[i]);
                }
            Add(new Izdel { Name = "Изделие11", Price = 50 }, new Link { IzdelId = 11, IzdelUp = 4, Kol = 2 });
        }
    }

    class Order
    {
        private List<Product> products = new List<Product>();

        private static Application ObjExcel = new Application();
        private static Workbook ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
        private static Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
       

        private static int row = 2;

        private void initOrder()
        {
            using (FIScontext db = new FIScontext())
            {
                var productes = db.Izdels.Join(db.Links,
                    I => I.Id,
                    L => L.IzdelId,
                    (I, L) => new Product
                    {
                        Id = I.Id,
                        Name = I.Name,
                        Parent = L.IzdelUp,
                        Quantity = L.Kol,
                        Price = L.Kol * I.Price,
                        ChildQuantity = 0,
                        Label = false
                    });
                foreach (var p in productes)
                {
                    products.Add(p);
                }
            }
            for(int i = 0; i < products.Count(); ++i)
            {
                for(int j = 0; j < products.Count(); ++j)
                {
                    if(products[i].Id == products[j].Parent && i != j)
                    {
                        products[i].children.Add(products[j]);
                        products[i].ChildQuantity++;
                    }
                }
            }
        }

        private Product FindById(long Id)
        {
            foreach(var p in products)
            {
                if (p.Id == Id) return p;
            }
            return null; 
        }

        public void deleteChildrenById(long Id)
        {
            Product parent = FindById(FindById(Id).Parent);
            if (parent != null)
            {
                foreach (var child in parent.children)
                {
                    if (child.Id == Id)
                    {
                        parent.children.Remove(child);
                        break;
                    }
                }
            }
        }

        private int exitLoop()
        {
            int exit = 1;
            foreach(var p in products)
            {
                if(p.Parent == -1)
                {
                    if (p.ChildQuantity == 0)
                    {
                        exit = 1;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            return exit;
        }

        private void calculatePrice()
        {
            while (exitLoop() == 0)
            {
                bool exit = false;
                while (!exit)
                {
                    foreach (var p in products)
                    {
                        if (p.ChildQuantity == 0 && p.Label == false)
                        {
                            Product parent = FindById(p.Parent);
                            if (parent != null)
                            {
                                parent.Price += p.Price;
                                parent.ChildQuantity--;
                                //deleteChildrenById(p.Id);
                                p.Label = true;
                                exit = false;
                            }
                        }
                    }
                    exit = true;
                }
            }
        }
       
        private void PrintExcel(string Indent, Product product)
        {
            try
            {
                ObjWorkSheet.Columns.AutoFit();

                ObjWorkSheet.Cells[1, 1] = "Изделие";
                ObjWorkSheet.Cells[1, 2] = "Количество";
                ObjWorkSheet.Cells[1, 3] = "Стоимость";
                
                ObjWorkSheet.Cells[row, 1] = Indent+ product.Name;
                ObjWorkSheet.Cells[row, 2] = product.Quantity;
                ObjWorkSheet.Cells[row, 3] = product.Price;
                ObjExcel.Visible = true;
                ObjExcel.UserControl = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка\n" + ex.Message);
            }
        }

        private void PrintPreety(Product product, string indent, bool last, int layer)
        {
            if (layer > 0)
            {
                layer--;
                PrintExcel(indent, product);
                row++;
                indent += "   ";
                for (int i = 0; i < product.children.Count; i++)
                {
                    PrintPreety(product.children[i], indent, i == product.children.Count - 1,layer);
                }
                
            }
        }


        public void printProducts()
        {
            initOrder();

            calculatePrice();

            foreach (var p in products)
            {
                if(p.Parent == -1)
                    PrintPreety(p, "", true,3);
            }
            
        }
    }

    class Program
    {
        
        static void Main(string[] args)
        {
            /*Инициализация базы данных */
            Database.initDatabase();
            
           // Database.PrintIzdel();
           // Console.WriteLine();

            Order order = new Order();
            /*Вывод отчета*/
            order.printProducts();
            Console.ReadLine();
        }
    }
}
